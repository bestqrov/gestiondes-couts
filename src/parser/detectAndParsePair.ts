import { parseLiquidation, type LiquidationResult } from './liquidation/liquidationParser.js';
import { parseDum, type DumResult } from './dum/dumParser.js';

export interface DetectedPair {
  liquidation: LiquidationResult;
  dum: DumResult;
  swapped: boolean;
}

// The user can upload the two documents into either slot in any order — try
// both assignments and use whichever one actually parses as a valid
// (Liquidation, DUM) pair, rather than trusting which upload field a file
// was placed in.
export function detectAndParsePair(textA: string, textB: string): DetectedPair {
  try {
    return { liquidation: parseLiquidation(textA), dum: parseDum(textB), swapped: false };
  } catch (firstAttemptError) {
    try {
      return { liquidation: parseLiquidation(textB), dum: parseDum(textA), swapped: true };
    } catch (secondAttemptError) {
      const firstMessage =
        firstAttemptError instanceof Error ? firstAttemptError.message : String(firstAttemptError);
      const secondMessage =
        secondAttemptError instanceof Error ? secondAttemptError.message : String(secondAttemptError);
      throw new Error(
        `Could not identify which uploaded file is the Liquidation and which is the DUM.\n` +
          `As (file 1 = Liquidation, file 2 = DUM): ${firstMessage}\n` +
          `As (file 2 = Liquidation, file 1 = DUM): ${secondMessage}`
      );
    }
  }
}
