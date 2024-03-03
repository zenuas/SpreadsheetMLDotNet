using Mina.Attributes;

namespace SpreadsheetMLDotNet.Data.Worksheets;

public enum ErrorValues
{
    [Alias("#DIV/0!")]
    DIV_0,

    [Alias("#GETTING_DATA")]
    GETTING_DATA,

    [Alias("#N/A")]
    NA,

    [Alias("#NAME?")]
    NAME,

    [Alias("#NULL!")]
    NULL,

    [Alias("#NUM!")]
    NUM,

    [Alias("#REF!")]
    REF,

    [Alias("#VALUE!")]
    VALUE,
}
