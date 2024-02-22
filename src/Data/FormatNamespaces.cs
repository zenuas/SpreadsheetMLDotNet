namespace SpreadsheetMLDotNet.Data;

public class FormatNamespaces
{
    public const string StrictURIBase = "http://purl.oclc.org/ooxml";
    public const string TransitionalURIBase = "http://schemas.openxmlformats.org";

    public const string StrictRelationship = $"{StrictURIBase}/officeDocument/relationships";
    public const string TransitionalRelationship = $"{TransitionalURIBase}/officeDocument/2006/relationships";

    public const string StrictOfficeDocument = $"{StrictRelationship}/officeDocument";
    public const string TransitionalOfficeDocument = $"{TransitionalRelationship}/officeDocument";

    public const string StrictSpreadsheetMLMain = $"{StrictURIBase}/spreadsheetml/main";
    public const string TransitionalSpreadsheetMLMain = $"{TransitionalURIBase}/spreadsheetml/2006/main";

    public const string StrictWorksheet = $"{StrictRelationship}/worksheet";
    public const string TransitionalWorksheet = $"{TransitionalRelationship}/worksheet";

    public const string StrictStyles = $"{StrictRelationship}/styles";
    public const string TransitionalStyles = $"{TransitionalRelationship}/styles";

    public const string StrictSharedStrings = $"{StrictRelationship}/sharedStrings";
    public const string TransitionalSharedStrings = $"{TransitionalRelationship}/sharedStrings";

    public static readonly string[] Relationships = [StrictRelationship, TransitionalRelationship];
    public static readonly string[] OfficeDocuments = [StrictOfficeDocument, TransitionalOfficeDocument];
    public static readonly string[] SpreadsheetMLMains = [StrictSpreadsheetMLMain, TransitionalSpreadsheetMLMain];
    public static readonly string[] Worksheets = [StrictWorksheet, TransitionalWorksheet];
    public static readonly string[] Styles = [StrictStyles, TransitionalStyles];
    public static readonly string[] SharedStrings = [StrictSharedStrings, TransitionalSharedStrings];
}
