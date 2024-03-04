﻿using System.Collections.Generic;

namespace SpreadsheetMLDotNet.Data.Styles;

public class CellStyles : IRelationshipable
{
    public List<Font> Fonts { get; init; } = [];
    public List<Fill> Fills { get; init; } = [];
    public List<Border> Borders { get; init; } = [];
}