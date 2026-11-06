export type TemplateErrorCode =
  // template mistakes
  | 'unclosed_tag'
  | 'unopened_tag'
  | 'unclosed_section'
  | 'unmatched_close'
  | 'invalid_expression'
  | 'unknown_filter'
  | 'filter_failed'
  | 'missing_value'
  | 'invalid_section'
  | 'invalid_image'
  | 'unsupported'
  // engine level
  | 'invalid_request'
  | 'invalid_options'
  | 'invalid_document'
  | 'unsupported_format'
  | 'invalid_model'
  | 'unknown_template'
  | 'limit_exceeded'
  | 'host_error'
  | 'internal'

/** One problem reported by the engine. */
export interface TemplateErrorInfo {
  code: TemplateErrorCode
  message: string
  /** Package part, e.g. `word/document.xml`. */
  part?: string
  /** Raw tag text, e.g. `{{#items}}`. */
  tag?: string
  /** Where in the part: `{ paragraph }`, `{ cell: 'B12' }`, `{ slide }`, … */
  location?: Record<string, unknown>
  /** Extra data, e.g. `{ hostMessage }` for failed filters or the exceeded limit. */
  details?: Record<string, unknown>
}

export class TemplateError extends Error {
  readonly code: TemplateErrorCode
  readonly part?: string
  readonly tag?: string
  readonly location?: Record<string, unknown>
  readonly details?: Record<string, unknown>
  /** All problems found. Compilation reports every template mistake at once; this error is the first. */
  readonly errors: TemplateErrorInfo[]

  constructor(message: string, info: Omit<TemplateErrorInfo, 'message'>, errors?: TemplateErrorInfo[]) {
    const where = [info.tag && `tag ${info.tag}`, info.part && `in ${info.part}`].filter(Boolean).join(' ')
    const more = errors && errors.length > 1 ? ` (+${errors.length - 1} more)` : ''
    super((where ? `${message} (${where})` : message) + more)
    this.name = 'TemplateError'
    this.code = info.code
    this.part = info.part
    this.tag = info.tag
    this.location = info.location
    this.details = info.details
    this.errors = errors ?? [{ message, ...info }]
  }

  static fromList(errors: TemplateErrorInfo[]): TemplateError {
    const [first] = errors
    return new TemplateError(first.message, first, errors)
  }
}
