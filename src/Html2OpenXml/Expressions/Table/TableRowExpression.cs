/* Copyright (C) Olivier Nizet https://github.com/onizet/html2openxml - All Rights Reserved
 * 
 * This source is subject to the Microsoft Permissive License.
 * Please see the License.txt file for more information.
 * All other rights reserved.
 * 
 * THIS CODE AND INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY 
 * KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE
 * IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A
 * PARTICULAR PURPOSE.
 */
using AngleSharp.Html.Dom;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace HtmlToOpenXml.Expressions;

/// <summary>
/// Process the parsing of a <c>tr</c> element which represent a row in a table.
/// </summary>
sealed class TableRowExpression : TableElementExpressionBase
{
    private readonly IHtmlTableRowElement rowNode;
    private readonly TableRowProperties rowProperties = new();
    private readonly int columCount;
    private readonly RowSpanCollection carriedRowSpans, rowSpans = [];


    public TableRowExpression(IHtmlTableRowElement node, int columCount, RowSpanCollection carriedRowSpans)
        : base (node)
    {
        rowNode = node;
        this.columCount = columCount;
        this.carriedRowSpans = carriedRowSpans;
    }

    /// <inheritdoc/>
    public override IEnumerable<OpenXmlElement> Interpret (ParsingContext context)
    {
        ComposeStyles(context);

        var cells = rowNode.Cells;
        if (cells.Length == 0 && !carriedRowSpans.Columns.Any())
            return [];

        var rowContext = context.CreateChild(this);
        var tableRow = new TableRow(rowProperties);
        int colIndex = 0;
        int cellIndex = 0;

        while (true)
        {
            if (carriedRowSpans.Contains(colIndex))
            {
                // this grid column is still occupied by a rowspan from a previous row
                int colSpan = carriedRowSpans.Decrement(colIndex);
                var mergedCell = TableCellExpression.CreateEmpty(
                    new VerticalMerge() { Val = MergedCellValues.Continue }
                );
                if (colSpan > 1) mergedCell.TableCellProperties!.GridSpan = new() { Val = colSpan };
                tableRow.AppendChild(mergedCell);

                colIndex += colSpan;
            }
            else if (cellIndex < cells.Length)
            {
                var cell = cells[cellIndex++];

                var expression = new TableCellExpression(cell);
                foreach (var element in expression.Interpret(rowContext))
                {
                    rowContext.CascadeStyles(element);
                    tableRow.AppendChild(element);
                }

                if (TableCellExpression.IsValidRowSpan(cell.RowSpan))
                    rowSpans.Add(colIndex, cell.RowSpan, cell.ColumnSpan);

                colIndex += cell.ColumnSpan;
            }
            else if (carriedRowSpans.Columns.Any(c => c > colIndex))
            {
                // gap before a carried span and no cell left to fill it
                tableRow.AppendChild(TableCellExpression.CreateEmpty());
                colIndex++;
            }
            else break;
        }

        // if the row is not complete, create empty cells
        if (colIndex < columCount)
            tableRow.AppendChild(TableCellExpression.CreateEmpty());

        rowSpans.UnionWith(carriedRowSpans);
        return [tableRow];
    }

    protected override void ComposeStyles(ParsingContext context)
    {
        base.ComposeStyles(context);

        Unit unit = styleAttributes.GetUnit("height", UnitMetric.Pixel);
        if (!unit.IsValid) unit = Unit.Parse(rowNode.GetAttribute("height").AsSpan(), UnitMetric.Pixel);

        switch (unit.Metric)
        {
            case UnitMetric.Point:
                rowProperties.AddChild(new TableRowHeight() { HeightType = HeightRuleValues.AtLeast, Val = (uint) (unit.Value * 20) });
                break;
            case UnitMetric.Pixel:
                rowProperties.AddChild(new TableRowHeight() { HeightType = HeightRuleValues.AtLeast, Val = (uint) unit.ValueInDxa });
                break;
        }
    }

    /// <summary>
    /// The carried row spans.
    /// </summary>
    public RowSpanCollection RowSpans
    {
        get => rowSpans;
    }
}
