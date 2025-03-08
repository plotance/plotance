// SPDX-FileCopyrightText: 2025 Plotance contributors <https://plotance.github.io/>
//
// SPDX-License-Identifier: MIT

using System.Globalization;
using Plotance.Models;
using Xunit;

namespace Plotance.Tests.Models;

public class LengthTest
{
    [Theory]
    [InlineData("10mm", 360000)]
    [InlineData("2.5cm", 900000)]
    [InlineData("1in", 914400)]
    [InlineData("72pt", 914400)]
    [InlineData("6pc", 914400)]
    [InlineData("96px", 914400)]
    [InlineData("0", 0)]
    [InlineData(" 0 ", 0)]
    [InlineData(" 10 mm ", 360000)]
    [InlineData("+10mm", 360000)]
    [InlineData("+0010mm", 360000)]
    [InlineData("0.5mm", 18000)]
    [InlineData("000.5mm", 18000)]
    [InlineData(".5mm", 18000)]
    [InlineData("-.5mm", -18000)]
    public void TextLength_WithValidInput_ReturnsCorrectEmuValue(
        string input,
        long expectedEmu
    )
    {
        // Arrange
        ILength length = new TextLength("test.md", 1, input);

        // Act
        var result = length.ToEmu();

        // Assert
        Assert.Equal(expectedEmu, result);
    }

    [Theory]
    [InlineData("")]
    [InlineData("10")]
    [InlineData("10m")]
    [InlineData("mm")]
    [InlineData("10 meters")]
    [InlineData("abc")]
    public void TextLength_WithInvalidInput_ThrowsPlotanceException(
        string input
    )
    {
        // Arrange
        ILength length = new TextLength("test.md", 1, input);

        // Act & Assert
        Assert.Throws<PlotanceException>(() => length.ToEmu());
    }

    [Fact]
    public void EmuLength_ReturnsCorrectValue()
    {
        // Arrange
        ILength length = new EmuLength(36000);

        // Act
        var result = length.ToEmu();

        // Assert
        Assert.Equal(36000, result);
    }

    [Fact]
    public void ILength_Zero_ReturnsZeroEmuLength()
    {
        // Arrange
        ILength zero = ILength.Zero;

        // Act
        var result = zero.ToEmu();

        // Assert
        Assert.Equal(0, result);
    }

    [Fact]
    public void ILength_FromPoint_ReturnsCorrectEmuLength()
    {
        // Arrange
        ILength length = ILength.FromPoint(72);

        // Act
        var result = length.ToEmu();

        // Assert
        Assert.Equal(914400, result); // 72pt = 1in = 914400 EMU
    }

    [Fact]
    public void ILength_FromEmu_ReturnsCorrectEmuLength()
    {
        // Arrange
        ILength length = ILength.FromEmu(360000);

        // Act
        var result = length.ToEmu();

        // Assert
        Assert.Equal(360000, result);
    }

    [Fact]
    public void TextLength_ToPoint_ReturnsCorrectValue()
    {
        // Arrange
        ILength length = new TextLength("test.md", 1, "72pt");

        // Act
        var result = length.ToPoint();

        // Assert
        Assert.Equal(72, result);
    }

    [Fact]
    public void TextLength_ToCentipoint_ReturnsCorrectValue()
    {
        // Arrange
        ILength length = new TextLength("test.md", 1, "1pt");

        // Act
        var result = length.ToCentipoint();

        // Assert
        Assert.Equal(100, result); // 1pt = 100 centipoints
    }
}
