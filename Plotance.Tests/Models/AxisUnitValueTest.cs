// SPDX-FileCopyrightText: 2025 Plotance contributors <https://plotance.github.io/>
//
// SPDX-License-Identifier: MIT

using System.Globalization;
using Plotance.Models;
using Xunit;

namespace Plotance.Tests.Models;

public class AxisUnitValueTest
{
    [Theory]
    [InlineData("5", 5)]
    [InlineData("5.5", 5.5)]
    [InlineData("10 days", 10)]
    [InlineData("1 day", 1)]
    [InlineData("24 hours", 1)]
    [InlineData("1 hour", 1.0 / 24)]
    [InlineData("30 minutes", 0.5 / 24)]
    [InlineData("30 minute", 0.5 / 24)]
    [InlineData("1 mins", 1.0 / 24 / 60)]
    [InlineData("1 min", 1.0 / 24 / 60)]
    [InlineData("30 seconds", 0.5 / 24 / 60)]
    [InlineData("30 second", 0.5 / 24 / 60)]
    [InlineData("1 secs", 1.0 / 24 / 60 / 60)]
    [InlineData("1 sec", 1.0 / 24 / 60 / 60)]
    [InlineData("1.5 hours", 1.5 / 24)]
    public void TextAxisUnitValue_WithValidInput_ReturnsCorrectDecimalValue(
        string input,
        double expectedValue
    )
    {
        // Arrange
        IAxisUnitValue unitValue = new TextAxisUnitValue("test.md", 1, input);

        // Act
        var result = unitValue.ToDouble();

        // Assert
        Assert.Equal(expectedValue, result, 10);
    }

    [Theory]
    [InlineData("")]
    [InlineData("-5")]
    [InlineData("abc")]
    [InlineData("5km")]
    [InlineData("5 meters")]
    [InlineData("5 months")]
    public void TextAxisUnitValue_WithInvalidInput_ThrowsPlotanceException(
        string input
    )
    {
        // Arrange
        IAxisUnitValue unitValue = new TextAxisUnitValue("test.md", 1, input);

        // Act & Assert
        Assert.Throws<PlotanceException>(() => unitValue.ToDecimal());
    }

    [Fact]
    public void DecimalAxisUnitValue_ReturnsCorrectValue()
    {
        // Arrange
        IAxisUnitValue unitValue = new DecimalAxisUnitValue(5.5m);

        // Act
        var result = unitValue.ToDecimal();

        // Assert
        Assert.Equal(5.5m, result);
    }

    [Fact]
    public void DecimalAxisUnitValue_ToDouble_ReturnsCorrectValue()
    {
        // Arrange
        IAxisUnitValue unitValue = new DecimalAxisUnitValue(5.5m);

        // Act
        var result = unitValue.ToDouble();

        // Assert
        Assert.Equal(5.5, result);
    }
}
