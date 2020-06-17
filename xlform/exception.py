class XlFormException(RuntimeError):
    """Base class for xlform exceptions"""

    pass


class XlFormRuntimeException(XlFormException):
    """Runtime exception"""

    pass


class XlFormNotImplementedException(XlFormRuntimeException):
    """Not implemented exception"""

    pass


class XlFormArgumentException(XlFormRuntimeException):
    """Argument exception"""

    pass


class XlFormValidationException(XlFormException):
    """Validation exception"""

    pass


class XlFormInternalException(XlFormRuntimeException):
    """Internal exception"""

    pass
