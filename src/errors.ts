export class OfficeMcpError extends Error {
  constructor(
    message: string,
    readonly code: string,
    readonly details?: Record<string, unknown>
  ) {
    super(message);
    this.name = "OfficeMcpError";
  }
}

export function toErrorResult(error: unknown) {
  const known = error instanceof OfficeMcpError;
  const message = error instanceof Error ? error.message : String(error);

  return {
    isError: true as const,
    content: [
      {
        type: "text" as const,
        text: JSON.stringify(
          {
            error: known ? error.code : "INTERNAL_ERROR",
            message,
            ...(known && error.details ? { details: error.details } : {})
          },
          null,
          2
        )
      }
    ]
  };
}
export function jsonResult(value: unknown) {
  return {
    content: [
      {
        type: "text" as const,
        text: JSON.stringify(value, null, 2)
      }
    ],
    structuredContent: value
  };
}
