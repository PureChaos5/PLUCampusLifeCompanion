/**
 * SWAP OPTIMIZATION PASS
 *
 * Consolidates an RA's fragmented shifts into longer contiguous blocks
 * matching their block preference. Every slot added to one shift is paid
 * for by removing a slot from another shift of the SAME RA on a different
 * day such that total hours never change for anyone.
 *
 * Two move types, always paired with a donate:
 *   CASE A — adjacent slot has room. Donate directly into it.
 *   CASE B — adjacent slot is full. Relocate one occupant to the opposite
 *            end of their own block (keeps them contiguous, keeps hours flat
 *            for them). Then donate into the now-open slot.
 *
 * Hard constraints always enforced:
 *   - One contiguous shift per day per RA
 *   - Unavailability for all RAs involved
 *   - maxRAsOnDesk never exceeded in any slot
 *   - Desk time boundaries
 */

function optimizeBlockPreferences(config, ras, schedules) {
  if (config.maxRAsOnDesk <= 1) return;

  Logger.log('=== SWAP OPTIMIZATION PASS ===');
  let totalMoves = 0;
  let passCount  = 0;
  let improved   = true;

  while (improved && passCount < 50) {
    improved = false;
    passCount++;

    for (const ra of ras) {
      const prefSlots = getMaxShift2(ra.blockPreference);

      // Sort longest first each pass so the biggest block always gets fed first.
      // This is critical: if two shifts are equal length, whichever goes first
      // will donate from the other. On the next pass the sort flips them (the
      // grown one is now longer) so it keeps getting fed. Without re-sorting
      // after every move they ping-pong and just shift around forever.
      const sorted = [...ra.assignments].sort((a, b) => b.length - a.length);

      let didMove = false;

      for (const assignment of sorted) {
        if (assignment.length >= prefSlots) continue;

        const building = assignment.building;
        const day      = assignment.day;
        const start    = assignment.startSlot;
        const len      = assignment.length;

        const donor = findDonorSlot(ra, assignment);
        if (!donor) continue;

        // Anti-pingpong guard: don't steal from a shift that's longer than ours.
        // We only consolidate INTO the longest (or equal) shift. After one move
        // the pass restarts and the sort ensures the grown shift stays on top.
        if (donor.assignment.length > assignment.length) continue;

        for (const direction of ['before', 'after']) {
          const targetSlot = direction === 'before' ? start - 1 : start + len;

          if (targetSlot < config.deskTimeStart || targetSlot >= config.deskTimeEnd) continue;
          if (ra.availability[day][targetSlot] === 'unavailable') continue;

          const targetOccupants = schedules[building][day][targetSlot];

          // ============================================================
          // CASE A: room at target — donate directly
          // ============================================================
          if (targetOccupants.length < config.maxRAsOnDesk) {
            executeDonate(ra, assignment, donor, targetSlot, direction, schedules);
            totalMoves++;
            improved = true;
            didMove = true;
            Logger.log(`  [CONSOLIDATE] ${ra.name}: grew ${direction} into slot ${targetSlot} (day ${day}, ${building}), donated from day ${donor.day} slot ${donor.slot}. Now ${assignment.length} slots.`);
            break;
          }

          // ============================================================
          // CASE B: target is full — relocate an occupant, then donate
          // ============================================================

          for (let oIdx = 0; oIdx < targetOccupants.length; oIdx++) {
            const otherName = targetOccupants[oIdx];
            if (otherName === ra.name) continue;
            const otherRA = ras.find(r => r.name === otherName);
            if (!otherRA) continue;

            const otherAssignment = otherRA.assignments.find(a =>
              a.building === building &&
              a.day      === day &&
              a.startSlot <= targetSlot &&
              (a.startSlot + a.length) > targetSlot
            );
            if (!otherAssignment) continue;

            // Only take from an edge
            const otherStart = otherAssignment.startSlot;
            const otherEnd   = otherStart + otherAssignment.length - 1;
            if (targetSlot !== otherStart && targetSlot !== otherEnd) continue;

            // Landing = opposite end of their block from what we took
            const landingSlot = (targetSlot === otherStart)
              ? otherEnd + 1
              : otherStart - 1;

            if (landingSlot < config.deskTimeStart || landingSlot >= config.deskTimeEnd) continue;
            if (otherRA.availability[day][landingSlot] === 'unavailable') continue;
            if (schedules[building][day][landingSlot].length >= config.maxRAsOnDesk) continue;

            // --- RELOCATE otherRA ---
            targetOccupants.splice(targetOccupants.indexOf(otherName), 1);
            schedules[building][day][landingSlot].push(otherName);

            if (targetSlot === otherStart) {
              otherAssignment.startSlot += 1;
            } else {
              otherAssignment.startSlot -= 1;
            }
            // length unchanged: lost one edge, gained the other

            // --- DONATE into the now-open targetSlot ---
            executeDonate(ra, assignment, donor, targetSlot, direction, schedules);

            totalMoves++;
            improved = true;
            didMove = true;
            Logger.log(`  [RELOCATE+CONSOLIDATE] ${ra.name}: moved ${otherName} from slot ${targetSlot} to ${landingSlot}, grew ${ra.name} ${direction} into ${targetSlot} (day ${day}, ${building}). Donated from day ${donor.day} slot ${donor.slot}.`);
            break;
          }

          if (didMove) break;
        }

        // After any successful move, break out of the assignment loop immediately.
        // The while loop will restart the pass, re-sorting assignments so the
        // now-longer shift gets priority next time.
        if (didMove) break;
      }

      // Break out of the RA loop too so the pass restarts cleanly.
      if (didMove) break;
    }
  }

  Logger.log(`=== SWAP OPTIMIZATION COMPLETE: ${totalMoves} moves in ${passCount} passes ===`);
}

/**
 * Execute a consolidation: remove one slot from the donor shift on another day,
 * add one slot to the current assignment at targetSlot. Hours flat.
 */
function executeDonate(ra, assignment, donor, targetSlot, direction, schedules) {
  // Remove donor slot from grid
  const donorOccupants = schedules[donor.building][donor.day][donor.slot];
  donorOccupants.splice(donorOccupants.indexOf(ra.name), 1);

  // Shrink or remove donor assignment
  if (donor.assignment.length <= 1) {
    ra.assignments.splice(ra.assignments.indexOf(donor.assignment), 1);
  } else if (donor.slot === donor.assignment.startSlot) {
    donor.assignment.startSlot += 1;
    donor.assignment.length   -= 1;
  } else {
    donor.assignment.length -= 1;
  }

  // Add RA to targetSlot and grow current assignment
  schedules[assignment.building][assignment.day][targetSlot].push(ra.name);
  if (direction === 'before') {
    assignment.startSlot -= 1;
  }
  assignment.length += 1;
}

/**
 * Find a donor slot: the end-slot of this RA's shortest assignment
 * on a different day. Returns { assignment, day, building, slot } or null.
 */
function findDonorSlot(ra, currentAssignment) {
  let bestDonor = null;

  for (const a of ra.assignments) {
    if (a === currentAssignment) continue;
    if (a.day === currentAssignment.day) continue;

    if (!bestDonor || a.length < bestDonor.length) {
      bestDonor = a;
    }
  }

  if (!bestDonor) return null;

  return {
    assignment: bestDonor,
    day:        bestDonor.day,
    building:   bestDonor.building,
    slot:       bestDonor.startSlot + bestDonor.length - 1
  };
}