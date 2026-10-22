export class OperationTimeoutError extends Error {
  constructor(message, code = "OPERATION_TIMEOUT") {
    super(message);
    this.name = "OperationTimeoutError";
    this.code = code;
  }
}

export function delay(milliseconds) {
  return new Promise((resolve) => globalThis.setTimeout(resolve, milliseconds));
}

export async function withTimeout(operation, milliseconds, message, code) {
  let timer;
  try {
    return await Promise.race([
      Promise.resolve(operation),
      new Promise((_, reject) => {
        timer = globalThis.setTimeout(
          () => reject(new OperationTimeoutError(message, code)),
          milliseconds,
        );
      }),
    ]);
  } finally {
    globalThis.clearTimeout(timer);
  }
}

export async function retry(operation, options = {}) {
  const attempts = options.attempts ?? 2;
  const delayMilliseconds = options.delayMilliseconds ?? 600;
  let lastError;

  for (let attempt = 1; attempt <= attempts; attempt += 1) {
    try {
      return await operation(attempt);
    } catch (error) {
      lastError = error;
      if (attempt >= attempts) break;
      options.onRetry?.(error, attempt + 1);
      await delay(delayMilliseconds);
    }
  }

  throw lastError;
}
