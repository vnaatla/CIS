namespace Cis.Api.Common;

/// <summary>
/// A request the user can fix (a rule of the business was broken). The message is shown to the user as is (the desktop's
/// MsgBox text when there is one); <see cref="ApiExceptionHandler"/> turns it into HTTP 400.
/// </summary>
public class BusinessRuleException(string message) : Exception(message);

/// <summary>The record asked for does not exist (HTTP 404).</summary>
public class NotFoundException(string message) : BusinessRuleException(message);

/// <summary>The request clashes with the current data, for example adding a record that already exists (HTTP 409).</summary>
public class ConflictException(string message) : BusinessRuleException(message);

/// <summary>The caller is not who they claim to be (wrong user name or password): HTTP 401.</summary>
public class UnauthorizedException(string message) : BusinessRuleException(message);

/// <summary>A feature needs a setting or an external system that is not configured on this server (HTTP 503).</summary>
public class NotConfiguredException(string message) : Exception(message);
