// =============================================================================
// srv/lib/calculations.js
// MIGP Sales Tool — Business Logic & Calculation Engine
//
// Implements all formulas from the functional spec:
//   - Size of Opportunity
//   - Annual Subscribed Load (with weighted average for escalations)
//   - Contract Lifetime Value (with weighted average + REC blending)
//   - Opportunity Subscribed Load (with weighted average)
//   - Term calculation (calendar days → decimal years)
//   - REC Blending (MIGP MWh, REC MWh, MIGP %, REC %)
//   - Metered / Unmetered consumption rollup
// =============================================================================

'use strict';

/**
 * Calculate Term in decimal years from two dates.
 * Spec: "Non-editable, calculated from start and end dates (calendar days).
 *        Must accommodate 2 decimals (e.g. 6.25 years = 6 years and 3 months)"
 *
 * @param {Date|string} startDate
 * @param {Date|string} endDate
 * @returns {number} term in years rounded to 2 decimal places
 */
function calculateTermYears(startDate, endDate) {
  const start = new Date(startDate);
  const end   = new Date(endDate);
  const diffMs   = end.getTime() - start.getTime();
  const diffDays = diffMs / (1000 * 60 * 60 * 24);
  // Use 365.25 to account for leap years
  const termYears = diffDays / 365.25;
  return round(termYears, 2);
}

/**
 * Calculate Size of Opportunity.
 * Spec Field 3: = Gross Annual Usage × Estimated Subscription %
 *
 * @param {number} grossAnnualUsageMWh  — sum of ALL CAs (selected or not)
 * @param {number} subscriptionPct      — 1 to 100
 * @returns {number}
 */
function calculateSizeOfOpportunity(grossAnnualUsageMWh, subscriptionPct) {
  return round(grossAnnualUsageMWh * (subscriptionPct / 100), 4);
}

/**
 * Calculate Est. Usage (MWh) for a single Provider Contract line.
 * Spec Field 9: = selected CAs 12-month usage × subscriptionPercent / 100
 *
 * @param {number} selectedCAsTotalUsageMWh
 * @param {number} subscriptionPercent
 * @returns {number}
 */
function calculateEstUsageMWh(selectedCAsTotalUsageMWh, subscriptionPercent) {
  return round(selectedCAsTotalUsageMWh * (subscriptionPercent / 100), 4);
}

/**
 * Calculate REC Blending values for a single Provider Contract line.
 *
 * Spec Field 14 — REC MWh:
 *   = {Size of Opp × NTE / (REC Price – Sub Fee)} – {Sub Fee × Size of Opp / (REC Price – Sub Fee)}
 *   = Size of Opp × (NTE – Sub Fee) / (REC Price – Sub Fee)
 *
 * Spec Field 13 — MIGP MWh:
 *   = Size of Opportunity – REC MWh
 *
 * Spec Field 11 — MIGP (%):
 *   = (MIGP MWh / Size of Opportunity) × 100
 *
 * Spec Field 12 — REC (%):
 *   = (REC MWh / Size of Opportunity) × 100
 *
 * @param {number} sizeOfOpportunityMWh
 * @param {number} ntePrice
 * @param {number} subscriptionFee
 * @param {number} recPrice
 * @returns {{ recMWh, migpMWh, migpPercent, recPercent }}
 */
function calculateRECBlending(sizeOfOpportunityMWh, ntePrice, subscriptionFee, recPrice) {
  const divisor = recPrice - subscriptionFee;

  if (divisor === 0) {
    throw new Error('REC Blending calculation error: REC Price and Subscription Fee cannot be equal (division by zero).');
  }

  const recMWh  = round(sizeOfOpportunityMWh * (ntePrice - subscriptionFee) / divisor, 4);
  const migpMWh = round(sizeOfOpportunityMWh - recMWh, 4);

  const migpPercent = sizeOfOpportunityMWh > 0
    ? round((migpMWh / sizeOfOpportunityMWh) * 100, 4)
    : 0;

  const recPercent = sizeOfOpportunityMWh > 0
    ? round((recMWh / sizeOfOpportunityMWh) * 100, 4)
    : 0;

  return { recMWh, migpMWh, migpPercent, recPercent };
}

/**
 * Calculate Annual Subscribed Load using weighted average.
 *
 * Spec Field 5:
 *   = Sum of (12-month usage of selected CAs × Subscription %) per portfolio
 *   With weighted average: = Σ (periodYears/totalTermYears × MIGP_or_EstUsage MWh)
 *
 * Non-escalation (single line): = selected CAs usage × subscriptionPct / 100
 * Escalation (multiple lines):  = Σ (periodYears[i] / totalTermYears) × migpMWh[i]
 *
 * Spec note: "If REC Blending use MIGP (MWh), otherwise use Est. Usage (MWh)"
 *
 * @param {Array} lines — array of provider contract line items
 *   Each: { periodYears, migpMWh, estUsageMWh, isRECBlend }
 * @param {number} totalTermYears
 * @returns {number}
 */
function calculateAnnualSubscribedLoad(lines, totalTermYears) {
  if (!lines || lines.length === 0) return 0;

  // Single line — no weighted average needed
  if (lines.length === 1) {
    const line = lines[0];
    return round(line.isRECBlend ? line.migpMWh : line.estUsageMWh, 4);
  }

  // Multiple lines (escalations) — weighted average
  // Spec example: (4/7) × 39,385.71 + (3/7) × 58,101.27
  let annualSubscribedLoad = 0;
  for (const line of lines) {
    const weight = totalTermYears > 0 ? line.periodYears / totalTermYears : 0;
    const mwh    = line.isRECBlend ? line.migpMWh : line.estUsageMWh;
    annualSubscribedLoad += weight * mwh;
  }
  return round(annualSubscribedLoad, 4);
}

/**
 * Calculate Contract Lifetime Value.
 *
 * Spec Field 8:
 *   Non-REC: = Annual Subscribed Load × Subscription Fee × Contract Term
 *   REC:     = MIGP MWh (calculated) × Subscription Fee × Contract Term
 *
 *   With weighted average for escalations:
 *   = Σ (subscriptionFee[i] × migpOrEstMWh[i] × periodYears[i])
 *
 * Spec example: (75 × 39,385.71 × 4) + (85 × 58,101.27 × 3)
 *
 * @param {Array} lines — provider contract lines
 *   Each: { subscriptionFee, periodYears, migpMWh, estUsageMWh, isRECBlend }
 * @returns {number}
 */
function calculateContractLifetimeValue(lines) {
  if (!lines || lines.length === 0) return 0;

  let clv = 0;
  for (const line of lines) {
    const mwh = line.isRECBlend ? line.migpMWh : line.estUsageMWh;
    clv += (line.subscriptionFee || 0) * mwh * (line.periodYears || 0);
  }
  return round(clv, 2);
}

/**
 * Calculate Opportunity Subscribed Load.
 *
 * Spec Field 9:
 *   = Annual Subscribed Load × Contract Term
 *
 *   With weighted average for escalations:
 *   = Σ (migpOrEstMWh[i] × periodYears[i])
 *
 * Spec example: 4 × 39,385.71 + 3 × 58,101.27
 *
 * @param {Array} lines — provider contract lines
 *   Each: { periodYears, migpMWh, estUsageMWh, isRECBlend }
 * @returns {number}
 */
function calculateOpportunitySubscribedLoad(lines) {
  if (!lines || lines.length === 0) return 0;

  let osl = 0;
  for (const line of lines) {
    const mwh = line.isRECBlend ? line.migpMWh : line.estUsageMWh;
    osl += mwh * (line.periodYears || 0);
  }
  return round(osl, 4);
}

/**
 * Roll up Gross Annual Usage from ALL contract accounts under all BPs
 * (regardless of whether CAs are selected or not — per spec Field 1)
 *
 * @param {Array} allContractAccounts — all CAs on the opportunity
 * @returns {number}
 */
function calculateGrossAnnualUsage(allContractAccounts) {
  if (!allContractAccounts || allContractAccounts.length === 0) return 0;
  const total = allContractAccounts.reduce((sum, ca) => sum + (ca.twelveMonthUsageMWh || 0), 0);
  return round(total, 4);
}

/**
 * Calculate metered and unmetered consumption from SELECTED CAs only.
 *
 * @param {Array} selectedContractAccounts — CAs where isSelected = true
 * @returns {{ meteredMWh, unmeteredMWh }}
 */
function calculateMeteredUnmetered(selectedContractAccounts) {
  let meteredMWh   = 0;
  let unmeteredMWh = 0;

  for (const ca of selectedContractAccounts || []) {
    if (ca.isMetered) {
      meteredMWh   += ca.twelveMonthUsageMWh || 0;
    } else {
      unmeteredMWh += ca.twelveMonthUsageMWh || 0;
    }
  }

  return {
    meteredMWh:   round(meteredMWh, 4),
    unmeteredMWh: round(unmeteredMWh, 4)
  };
}

/**
 * Master calculation function — orchestrates all calculations for a header.
 * Called by the Calculate action handler.
 *
 * @param {Object} header           — header record
 * @param {Array}  allCAs           — all contract accounts (for gross usage)
 * @param {Array}  selectedCAs      — selected contract accounts
 * @param {Array}  providerContracts — product/pricing lines
 * @param {boolean} isRECBlend      — true if header.priceStructure = REC_BLEND
 * @returns {Object} calculated field values to persist back to DB
 */
function runFullCalculation(header, allCAs, selectedCAs, providerContracts, isRECBlend) {
  // --- Step 1: Gross Annual Usage (all CAs, regardless of selection) ---
  const grossAnnualUsageMWh = calculateGrossAnnualUsage(allCAs);

  // --- Step 2: Size of Opportunity ---
  const sizeOfOpportunityMWh = calculateSizeOfOpportunity(
    grossAnnualUsageMWh,
    header.estimatedSubscriptionPct
  );

  // --- Step 3: Term (from header dates, or sum of line item dates) ---
  let totalTermYears = 0;
  if (header.termStartDate && header.termEndDate) {
    totalTermYears = calculateTermYears(header.termStartDate, header.termEndDate);
  }

  // --- Step 4: Per-line calculations (Est. Usage, REC Blending, Period Years) ---
  const selectedCATotalUsage = selectedCAs.reduce(
    (sum, ca) => sum + (ca.twelveMonthUsageMWh || 0), 0
  );

  const enrichedLines = [];
  let derivedTermYears = 0;

  for (const pc of providerContracts || []) {
    // Period years for this escalation segment
    const periodYears = (pc.startDate && pc.endDate)
      ? calculateTermYears(pc.startDate, pc.endDate)
      : 0;
    derivedTermYears += periodYears;

    // Est. Usage for this line
    const estUsageMWh = calculateEstUsageMWh(
      selectedCATotalUsage,
      pc.subscriptionPercent || 0
    );

    // REC Blending (LCVP only, when price structure = REC_BLEND)
    let recMWh = 0, migpMWh = 0, migpPercent = 0, recPercent = 0;
    if (isRECBlend && header.ntePrice && pc.subscriptionFee && pc.recPrice) {
      ({ recMWh, migpMWh, migpPercent, recPercent } = calculateRECBlending(
        sizeOfOpportunityMWh,
        header.ntePrice,
        pc.subscriptionFee,
        pc.recPrice
      ));
    } else {
      // Non-REC: MIGP MWh = Est. Usage (full subscription load)
      migpMWh = estUsageMWh;
    }

    enrichedLines.push({
      id: pc.id,
      subscriptionFee: pc.subscriptionFee || 0,
      periodYears,
      estUsageMWh,
      migpMWh,
      recMWh,
      migpPercent,
      recPercent,
      isRECBlend
    });
  }

  // Use derived term if header dates not set
  if (totalTermYears === 0) totalTermYears = round(derivedTermYears, 2);

  // --- Step 5: Annual Subscribed Load (weighted average) ---
  const annualSubscribedLoadMWh = calculateAnnualSubscribedLoad(enrichedLines, totalTermYears);

  // --- Step 6: Contract Lifetime Value ---
  const contractLifetimeValue = calculateContractLifetimeValue(enrichedLines);

  // --- Step 7: Opportunity Subscribed Load ---
  const opportunitySubscribedLoad = calculateOpportunitySubscribedLoad(enrichedLines);

  // --- Step 8: Metered / Unmetered breakdown ---
  const { meteredMWh, unmeteredMWh } = calculateMeteredUnmetered(selectedCAs);

  return {
    // Header-level results
    grossAnnualUsageMWh,
    sizeOfOpportunityMWh,
    termYears: totalTermYears,
    annualSubscribedLoadMWh,
    contractLifetimeValue,
    opportunitySubscribedLoad,
    meteredEstConsumptionMWh:   meteredMWh,
    unmeteredEstConsumptionMWh: unmeteredMWh,

    // Per-line results (to update ProviderContract rows)
    lineItems: enrichedLines
  };
}

/**
 * Validate business rules before allowing save/calculate.
 * Returns array of validation errors (empty = OK).
 *
 * @param {Object} header
 * @param {Array}  providerContracts
 * @param {string} category  — SMB / LCVP / DEDICATED_ARRAY
 * @returns {string[]} errors
 */
function validateBeforeCalculate(header, providerContracts, category) {
  const errors = [];
  const isLCVP = category === 'LCVP' || category === 'DEDICATED_ARRAY';

  for (const pc of providerContracts || []) {
    // Spec Exception 030: Required field missing
    if (!pc.portfolio) {
      errors.push(`Line ${pc.lineItemNumber}: Portfolio is required.`);
    }
    if (!pc.startDate) {
      errors.push(`Line ${pc.lineItemNumber}: Start Date is required.`);
    }
    if (!pc.subscriptionFee && !pc.fixedPrice && !pc.fixedMWh) {
      errors.push(`Line ${pc.lineItemNumber}: Subscription Fee, Fixed Price, or Fixed MWh is required.`);
    }

    // SMB: must enter either Fixed Price OR Subscription % — not both
    if (!isLCVP) {
      if (pc.fixedPrice && pc.subscriptionPercent) {
        errors.push(`Line ${pc.lineItemNumber}: SMB — enter either Fixed Price or Subscription %, not both.`);
      }
      // SMB: must not have end date (no escalations)
      if (providerContracts.length > 1) {
        errors.push('SMB opportunities must have only 1 product line item (no escalations).');
        break; // One error is enough
      }
    }

    // LCVP: must enter either Fixed MWh OR Subscription % — not both
    if (isLCVP) {
      if (pc.fixedMWh && pc.subscriptionPercent) {
        errors.push(`Line ${pc.lineItemNumber}: LCVP — enter either Fixed MWh or Subscription %, not both.`);
      }
    }
  }

  // REC Blending requires NTE Price
  if (header.priceStructure === 'REC_BLEND' && !header.ntePrice) {
    errors.push('NTE Price is required when Price Structure is REC Blend.');
  }

  return errors;
}

// ---------------------------------------------------------------------------
// Utility
// ---------------------------------------------------------------------------
function round(value, decimals) {
  const factor = Math.pow(10, decimals);
  return Math.round((value + Number.EPSILON) * factor) / factor;
}

module.exports = {
  calculateTermYears,
  calculateSizeOfOpportunity,
  calculateEstUsageMWh,
  calculateRECBlending,
  calculateAnnualSubscribedLoad,
  calculateContractLifetimeValue,
  calculateOpportunitySubscribedLoad,
  calculateGrossAnnualUsage,
  calculateMeteredUnmetered,
  runFullCalculation,
  validateBeforeCalculate,
  round
};
