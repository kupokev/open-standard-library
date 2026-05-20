/// <summary>
/// Specifies the character encoding used when reading or writing delimited files (CSV, TSV, etc.).
/// Does not apply to ODS or XLSX formats, which require UTF-8 per their specifications.
/// </summary>
public enum FileEncoding
{
    /// <summary>
    /// UTF-8 encoding (default)
    /// </summary>
    UTF8 = 0,

    /// <summary>
    /// ASCII encoding (7-bit)
    /// </summary>
    ASCII = 1,

    /// <summary>
    /// Unicode encoding (UTF-16 little-endian)
    /// </summary>
    Unicode = 2,

    /// <summary>
    /// UTF-32 encoding (little-endian)
    /// </summary>
    UTF32 = 3
}
