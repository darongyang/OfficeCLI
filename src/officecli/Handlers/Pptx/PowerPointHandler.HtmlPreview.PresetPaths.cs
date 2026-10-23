// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.Globalization;
using DocumentFormat.OpenXml;
using Drawing = DocumentFormat.OpenXml.Drawing;

namespace OfficeCli.Handlers;

public partial class PowerPointHandler
{
    /// <summary>
    /// DrawingML leftBrace/rightBrace are open stroked paths, not rectangular
    /// borders or closed clip-path polygons. Coordinates are in the shape's EMU
    /// view box; the surrounding shape retains its rotation, flips and grouping.
    /// </summary>
    private static string PresetBraceToSvgPath(string preset, Drawing.PresetGeometry geometry,
        long width, long height)
    {
        if (width <= 0 || height <= 0) return "";
        double w = width, h = height, shortSide = Math.Min(w, h);
        var a2 = Math.Clamp(ReadBraceAdjustment(geometry, "adj2", 50000), 0, 100000);
        var maxA1 = Math.Min(a2, 100000 - a2) / 2 * h / shortSide;
        var a1 = Math.Clamp(ReadBraceAdjustment(geometry, "adj1", 8333), 0, maxA1);
        var radiusX = w / 2;
        var radiusY = shortSide * a1 / 100000;
        var middleY = h * a2 / 100000;
        var right = preset == "rightBrace";
        var endX = right ? 0 : w;
        var middleX = right ? w : 0;
        var outerSweep = right ? 0 : 1;
        var innerSweep = right ? 1 : 0;
        return $"M {endX:0.##} {h:0.##} " +
            $"A {radiusX:0.##} {radiusY:0.##} 0 0 {outerSweep} {radiusX:0.##} {h - radiusY:0.##} " +
            $"L {radiusX:0.##} {middleY + radiusY:0.##} " +
            $"A {radiusX:0.##} {radiusY:0.##} 0 0 {innerSweep} {middleX:0.##} {middleY:0.##} " +
            $"A {radiusX:0.##} {radiusY:0.##} 0 0 {innerSweep} {radiusX:0.##} {middleY - radiusY:0.##} " +
            $"L {radiusX:0.##} {radiusY:0.##} " +
            $"A {radiusX:0.##} {radiusY:0.##} 0 0 {outerSweep} {endX:0.##} 0";
    }

    private static double ReadBraceAdjustment(Drawing.PresetGeometry geometry, string name, double fallback)
    {
        var formula = geometry.GetFirstChild<Drawing.AdjustValueList>()?
            .Elements<Drawing.ShapeGuide>().FirstOrDefault(g => g.Name?.Value == name)?.Formula?.Value;
        var parts = formula?.Split(' ', StringSplitOptions.RemoveEmptyEntries);
        return parts is { Length: 2 } && parts[0] == "val"
            && double.TryParse(parts[1], NumberStyles.Float, CultureInfo.InvariantCulture, out var value)
            && double.IsFinite(value) ? value : fallback;
    }
}
