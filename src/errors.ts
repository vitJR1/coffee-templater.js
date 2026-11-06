export type TemplateErrorCode =
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

export interface TemplateErrorDetails {
  code: TemplateErrorCode
  /** Part of the package where the error happened, e.g. `word/document.xml`. */
  part?: string
  /** The raw tag text, e.g. `{{#items}}`. */
  tag?: string
}

export class TemplateError extends Error {
  readonly code: TemplateErrorCode
  readonly part?: string
  readonly tag?: string

  constructor(message: string, details: TemplateErrorDetails) {
    const where = [details.tag && `tag ${details.tag}`, details.part && `in ${details.part}`].filter(Boolean).join(' ')
    super(where ? `${message} (${where})` : message)
    this.name = 'TemplateError'
    this.code = details.code
    this.part = details.part
    this.tag = details.tag
  }
}
