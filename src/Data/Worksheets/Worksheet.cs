﻿using Mina.Extension;
using SpreadsheetMLDotNet.Data.Address;
using SpreadsheetMLDotNet.Data.AutoFilterSettings;
using System;
using System.Collections.Generic;
using System.Linq;

namespace SpreadsheetMLDotNet.Data.Worksheets;

public class Worksheet : IRelationshipable
{
    public required string Name { get; set; }
    public IndexedList<Row> Rows { get; init; } = new() { New = () => new() };
    public IndexedList<Column> Columns { get; init; } = new() { New = () => new() };
    public List<IAddressRange> Merges { get; init; } = [];
    public AutoFilter? AutoFilter { get; set; }
    public SheetStates? SheetState { get; set; }


    public Row? GetRowOrDefault(int row) => Rows.GetOrDefault(row);

    public Row GetRow(int row) => Rows.GetOrNewAdd(row);

    public IEnumerable<(int Index, Row Row)> GetRowsExists(string range) => SpreadsheetML.ConvertRowRange(range).To(x => GetRowsExists(x.From, x.To));

    public IEnumerable<(int Index, Row Row)> GetRowsExists(int from, int to) => Lists.RangeTo(Math.Min(from, to), Math.Max(from, to)).Where(x => x >= Rows.StartIndex && x < Rows.StartIndex + Rows.Count).Select(x => (x, Rows[x - Rows.StartIndex]));

    public IEnumerable<(int Index, Row Row)> GetRows(string range) => SpreadsheetML.ConvertRowRange(range).To(x => GetRows(x.From, x.To));

    public IEnumerable<(int Index, Row Row)> GetRows(int from, int to) => Lists.RangeTo(Math.Min(from, to), Math.Max(from, to)).Select(x => (x, Rows.GetOrNewAdd(x)));

    public Column? GetColumnOrDefault(string column) => GetColumnOrDefault(SpreadsheetML.ConvertColumnNameToIndex(column));

    public Column? GetColumnOrDefault(int column) => Columns.GetOrDefault(column);

    public Column GetColumn(string column) => GetColumn(SpreadsheetML.ConvertColumnNameToIndex(column));

    public Column GetColumn(int column) => Columns.GetOrNewAdd(column);

    public IEnumerable<(int Index, Column Column)> GetColumnsExists(string range) => SpreadsheetML.ConvertColumnRange(range).To(x => GetColumnsExists(x.From, x.To));

    public IEnumerable<(int Index, Column Column)> GetColumnsExists(int from, int to) => Lists.RangeTo(Math.Min(from, to), Math.Max(from, to)).Where(x => x >= Columns.StartIndex && x < Columns.StartIndex + Columns.Count).Select(x => (x, Columns[x - Columns.StartIndex]));

    public IEnumerable<(int Index, Column Column)> GetColumns(string range) => SpreadsheetML.ConvertColumnRange(range).To(x => GetColumns(x.From, x.To));

    public IEnumerable<(int Index, Column Column)> GetColumns(int from, int to) => Lists.RangeTo(Math.Min(from, to), Math.Max(from, to)).Select(x => (x, Columns.GetOrNewAdd(x)));

    public Cell? GetCellOrDefault(string address) => SpreadsheetML.ConvertCellAddress(address).To(x => GetCellOrDefault(x.Row, x.Column));

    public Cell? GetCellOrDefault(int row, int column) => Rows.GetOrDefault(row)?.Cells?.GetOrDefault(column);

    public Cell GetCell(string address) => SpreadsheetML.ConvertCellAddress(address).To(x => GetCell(x.Row, x.Column));

    public Cell GetCell(int row, int column) => Rows.GetOrNewAdd(row).Cells.GetOrNewAdd(column);

    public void SetCell(string address, Cell cell) => SpreadsheetML.ConvertCellAddress(address).Return(x => SetCell(x.Row, x.Column, cell));

    public void SetCell(int row, int column, Cell cell) => Rows.GetOrNewAdd(row).Cells.SetValue(column, cell);

    public void MergeRows(RowAddressRange range) => Merges.Add(range);

    public void MergeColumns(ColumnAddressRange range) => Merges.Add(range);

    public void MergeCells(CellAddressRange range) => Merges.Add(range);

    public void Merge(IAddressRange range) => Merges.Add(range);
}
