namespace Cis.Api.Common;

/// <summary>SQL Server error numbers the API maps to user-facing responses.</summary>
public static class SqlErrorNumbers
{
    /// <summary>String or binary data would be truncated (pre-2019 message).</summary>
    public const int StringTruncated = 8152;

    /// <summary>String or binary data would be truncated in a named column (2019+ message).</summary>
    public const int StringTruncatedInColumn = 2628;

    /// <summary>Transaction chosen as the deadlock victim.</summary>
    public const int Deadlock = 1205;

    /// <summary>Command timeout raised by SqlClient.</summary>
    public const int Timeout = -2;

    /// <summary>Could not find stored procedure: the SQL scripts of this version have not been run.</summary>
    public const int MissingProcedure = 2812;

    /// <summary>Invalid object name: usually a synonym whose target database is missing.</summary>
    public const int InvalidObject = 208;

    /// <summary>Violation of a primary key or unique constraint.</summary>
    public const int DuplicateKey = 2627;

    /// <summary>Cannot insert a duplicate key row in a unique index.</summary>
    public const int DuplicateIndexKey = 2601;

    /// <summary>True for both truncation messages.</summary>
    public static bool IsTruncation(int number) => number is StringTruncated or StringTruncatedInColumn;

    /// <summary>True for both duplicate-key messages.</summary>
    public static bool IsDuplicateKey(int number) => number is DuplicateKey or DuplicateIndexKey;
}
