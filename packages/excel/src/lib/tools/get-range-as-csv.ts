import { Type } from "@sinclair/typebox";
import { getRangeAsCsv } from "../excel/api";
import { defineTool, toolError, toolSuccess } from "./types";

export const getRangeAsCsvTool = defineTool({
  name: "get_range_as_csv",
  label: "Get Range as CSV",
  description:
    "Read cell data from a range and return it as CSV format. " +
    "Great for analysis with pandas or quick data inspection. " +
    "Use this when you need tabular data without styling info.",
  parameters: Type.Object({
    sheetId: Type.Number({ description: "The worksheet ID (1-based index)" }),
    range: Type.String({ description: "Range in A1 notation, e.g. 'A1:Z100'" }),
    includeHeaders: Type.Optional(
      Type.Boolean({
        description: "Include first row as headers. Default: true",
        default: true,
      }),
    ),
    maxRows: Type.Optional(
      Type.Number({
        description: "Maximum rows to return. Default: 500",
        default: 500,
      }),
    ),
    explanation: Type.Optional(
      Type.String({
        description: "Brief explanation of what you're reading (max 255 chars)",
        maxLength: 255,
      }),
    ),
  }),
  execute: async (toolCallId, params) => {
    console.log(`[tool:get_range_as_csv] ${toolCallId} start`, params);
    try {
      const result = await getRangeAsCsv(params.sheetId, params.range, {
        includeHeaders: params.includeHeaders,
        maxRows: params.maxRows,
      });
      console.log(`[tool:get_range_as_csv] ${toolCallId} success`);
      return toolSuccess(result);
    } catch (error) {
      console.error(`[tool:get_range_as_csv] ${toolCallId} error`, error);
      const message =
        error instanceof Error ? error.message : "Unknown error reading CSV";
      return toolError(message);
    }
  },
});
