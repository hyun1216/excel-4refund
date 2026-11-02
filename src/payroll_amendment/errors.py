class PayrollAmendmentError(Exception):
    code = "E_UNKNOWN"


class ConfigError(PayrollAmendmentError):
    code = "E_CONFIG_SCHEMA"


class SourceNotFound(PayrollAmendmentError):
    code = "E_SOURCE_NOT_FOUND"


class SchemaMismatch(PayrollAmendmentError):
    code = "E_SCHEMA_MISMATCH"


class DuplicateEmployee(PayrollAmendmentError):
    code = "E_DUPLICATE_EMPLOYEE"


class DuplicatePdfIndex(PayrollAmendmentError):
    code = "E_DUPLICATE_PDF_INDEX"


class PdfMatchError(PayrollAmendmentError):
    code = "E_PDF_MATCH"


class PdfSplitError(PayrollAmendmentError):
    code = "E_PDF_SPLIT"


class PdfToExcelError(PayrollAmendmentError):
    code = "E_PDF_TO_EXCEL"


class RuleNotImplemented(PayrollAmendmentError):
    code = "E_RULE_NOT_IMPLEMENTED"

