UnknownReason = 'UnknownReason'


class GetIssueFailedError(Exception):
    def __init__(self, issue_key: str = None, issue_type: str = None, reason=UnknownReason):
        self.issue_info = "[%s(%s)]" % (issue_key, issue_type)
        super().__init__("Getting Issue: %s failed, because: %s." % (self.issue_info, reason))
        self.reason = reason


class GetParentFailedError(GetIssueFailedError):
    def __init__(self, issue_key: str = None, reason=UnknownReason):
        super().__init__(issue_key=issue_key, issue_type='Task', reason=reason)


class GetEpicFailedError(GetIssueFailedError):
    def __init__(self, issue_key: str = None, reason=UnknownReason):
        super().__init__(issue_key=issue_key, issue_type='Epic', reason=reason)


class CoordinateError(Exception):
    def __init__(self, reason=UnknownReason):
        super().__init__("Unable to generate coordinates, because: %s." % reason)
        self.reason = reason


class InvalidFieldError(CoordinateError):
    def __init__(self, attr: str = None):
        super().__init__("No Attribute: %s" % attr)


class BadEpicError(CoordinateError):
    def __init__(self):
        super().__init__("Bad Epic")


class MisMatchingError(Exception):
    def __init__(self, coord: tuple = None, reason=UnknownReason):
        super().__init__("Mismatching coordinate: %s, because: %s." % (coord, reason))


class NoMatchingError(MisMatchingError):
    def __init__(self, coord: tuple = None):
        super().__init__(coord, "No coordinates were matched")


class ManyMatchingError(MisMatchingError):
    def __init__(self, coord: tuple = None):
        super().__init__(coord, "Multiple coordinates matched")


class MatchingNAError(MisMatchingError):
    def __init__(self, coord: tuple = None):
        super().__init__(coord, "The coordinates match but the value is nan")
