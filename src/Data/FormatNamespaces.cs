namespace SpreadsheetMLDotNet.Data;

public class FormatNamespaces
{
    public const string StrictRelationship = "http://purl.oclc.org/ooxml/officeDocument/relationships";
    public const string TransitionalRelationship = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    public const string StrictOfficeDocument = $"{StrictRelationship}/officeDocument";
    public const string TransitionalOfficeDocument = $"{TransitionalRelationship}/officeDocument";

    public const string StrictWorksheet = $"{StrictRelationship}/worksheet";
    public const string TransitionalWorksheet = $"{TransitionalRelationship}/worksheet";

    public const string StrictStyles = $"{StrictRelationship}/styles";
    public const string TransitionalStyles = $"{TransitionalRelationship}/styles";

    public const string StrictSharedStrings = $"{StrictRelationship}/sharedStrings";
    public const string TransitionalSharedStrings = $"{TransitionalRelationship}/sharedStrings";

    public static readonly string[] Relationships = [StrictRelationship, TransitionalRelationship];
    public static readonly string[] OfficeDocuments = [StrictOfficeDocument, TransitionalOfficeDocument];
    public static readonly string[] Worksheets = [StrictWorksheet, TransitionalWorksheet];
    public static readonly string[] Styles = [StrictStyles, TransitionalStyles];
    public static readonly string[] SharedStrings = [StrictSharedStrings, TransitionalSharedStrings];
}
