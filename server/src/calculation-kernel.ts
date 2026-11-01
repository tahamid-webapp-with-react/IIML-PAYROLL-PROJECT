import Decimal from "decimal.js-light";

export type QuantityRate = { quantity: string | number; rate: string | number };
export type CustomTier = {
  id?: string;
  label?: string;
  minHours?: string | number | null;
  maxHours?: string | number | null;
  multiplier: string | number;
};
export type SlabConfiguration = {
  thresholdLow?: string | number;
  thresholdHigh?: string | number;
  multiplierLow?: string | number;
  multiplierMid?: string | number;
  multiplierHigh?: string | number;
  customTiers?: CustomTier[];
};

/**
 * Preserves the verified workbook kernel without binding it to a project.
 * Inputs must already be normalized, reconciled, and authorized before Stage 5
 * calls this function. The caller chooses the quantity/rate pairs and slab.
 */
export function calculateVerifiedKernel(
  quantityRates: QuantityRate[],
  averageHours: string | number,
  slab: SlabConfiguration,
) {
  const base = quantityRates.reduce(
    (sum, item) => sum.plus(new Decimal(item.quantity).times(item.rate)),
    new Decimal(0),
  );
  const hours = new Decimal(averageHours);
  let multiplier: Decimal;

  if (slab.customTiers && slab.customTiers.length > 0) {
    let matchedTier: CustomTier | undefined;
    for (const tier of slab.customTiers) {
      const minOk = tier.minHours == null || String(tier.minHours).trim() === "" || hours.gte(new Decimal(tier.minHours));
      const maxOk = tier.maxHours == null || String(tier.maxHours).trim() === "" || hours.lte(new Decimal(tier.maxHours));
      if (minOk && maxOk) {
        matchedTier = tier;
        break;
      }
    }
    if (matchedTier) {
      multiplier = new Decimal(matchedTier.multiplier);
    } else {
      multiplier = new Decimal(1);
    }
  } else {
    const low = new Decimal(slab.thresholdLow ?? 4);
    const high = new Decimal(slab.thresholdHigh ?? 6);
    multiplier = hours.lt(low)
      ? new Decimal(slab.multiplierLow ?? 0.6)
      : hours.lt(high)
        ? new Decimal(slab.multiplierMid ?? 0.8)
        : new Decimal(slab.multiplierHigh ?? 1.0);
  }

  return {
    baseAmount: base.toDecimalPlaces(6).toString(),
    multiplier: multiplier.toString(),
    payableAmount: base.times(multiplier).toDecimalPlaces(6).toString(),
  };
}
