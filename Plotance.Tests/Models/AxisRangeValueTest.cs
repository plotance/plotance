// SPDX-FileCopyrightText: 2025 Plotance contributors <https://plotance.github.io/>
//
// SPDX-License-Identifier: MIT

using System.Globalization;
using Plotance.Models;
using Xunit;

namespace Plotance.Tests.Models;

public class AxisRangeValueTest
{
    [Theory]
    [InlineData("auto", null)]
    [InlineData("10", 10.0)]
    [InlineData("-5.5", -5.5)]
    [InlineData("2025-01-01", 45658.0)]
    [InlineData("2025-01-01T10:00:00", 45658.4166666667)]
    [InlineData("2025-01-01 12:00:00", 45658.5)]
    [InlineData("2025-01-01 12:34:56", 45658.5242592592592593)]
    [InlineData("2025-01-01 12:34:56.789", 45658.5242683912037037)]
    [InlineData("00:00:00", 0.0)]
    [InlineData("12:00:00", 0.5)]
    [InlineData("12:34:56", 0.5242592592592593)]
    [InlineData("12:34:56.789", 0.5242683912037037)]
    public void TextAxisRangeValue_WithValidInput_ReturnsCorrectValue(
        string input,
        double? expectedValue
    )
    {
        // Arrange
        IAxisRangeValue rangeValue = new TextAxisRangeValue(
            "test.md",
            1,
            input
        );

        // Act
        var result = rangeValue.ToDouble();

        // Assert
        if (expectedValue == null)
        {
            Assert.Null(result);
        }
        else
        {
            Assert.NotNull(result);
            Assert.Equal(expectedValue.Value, result.Value, 10);
        }
    }

    [Theory]
    [InlineData("")]
    [InlineData("abc")]
    [InlineData("invalid")]
    [InlineData("2025/01/01")]
    [InlineData("01-02-2025")]
    [InlineData("12pm")]
    [InlineData("12:34:56 am")]
    public void TextAxisRangeValue_WithInvalidInput_ThrowsPlotanceException(
        string input
    )
    {
        // Arrange
        IAxisRangeValue rangeValue = new TextAxisRangeValue(
            "test.md",
            1,
            input
        );

        // Act & Assert
        Assert.Throws<PlotanceException>(() => rangeValue.ToDecimal());
    }

    [Fact]
    public void DecimalAxisRangeValue_ReturnsCorrectValue()
    {
        // Arrange
        IAxisRangeValue rangeValue = new DecimalAxisRangeValue(5.5m);

        // Act
        var result = rangeValue.ToDecimal();

        // Assert
        Assert.Equal(5.5m, result);
    }

    [Fact]
    public void DecimalAxisRangeValue_ToDouble_ReturnsCorrectValue()
    {
        // Arrange
        IAxisRangeValue rangeValue = new DecimalAxisRangeValue(5.5m);

        // Act
        var result = rangeValue.ToDouble();

        // Assert
        Assert.Equal(5.5, result);
    }
}
