using System;
using System.Collections.Generic;
using CodingWithCalvin.MCPServer.Services;
using CodingWithCalvin.MCPServer.Shared.Models;
using Xunit;

namespace CodingWithCalvin.MCPServer.Tests;

public class TestTargetSelectionTests
{
	[Fact]
	public void FindTestsForTarget_PrefersExactFullyQualifiedName()
	{
		var tests = new List<TestExplorerTestInfo>
		{
			Test("MyApp.Tests.CalculatorTests.Add_ReturnsSum", "Add_ReturnsSum"),
			Test("MyApp.Tests.CalculatorTests.Add_ReturnsDifference", "Add_ReturnsDifference")
		};

		var selected = TestExplorerInterop.FindTestsForTarget(
			tests,
			"MyApp.Tests.CalculatorTests.Add_ReturnsSum");

		Assert.Equal("MyApp.Tests.CalculatorTests.Add_ReturnsSum", Assert.Single(selected).FullyQualifiedName);
	}

	[Fact]
	public void FindTestsForTarget_SelectsAllTestsInFullyQualifiedClass()
	{
		var tests = new List<TestExplorerTestInfo>
		{
			Test("MyApp.Tests.CalculatorTests.Add_ReturnsSum", "Add_ReturnsSum"),
			Test("MyApp.Tests.CalculatorTests.Add_ReturnsDifference", "Add_ReturnsDifference"),
			Test("MyApp.Tests.StringTests.Trim", "Trim")
		};

		var selected = TestExplorerInterop.FindTestsForTarget(tests, "MyApp.Tests.CalculatorTests");

		Assert.Equal(2, selected.Count);
	}

	[Fact]
	public void FindTestsForTarget_ReturnsAllSimpleNameMatchesForAmbiguityHandling()
	{
		var tests = new List<TestExplorerTestInfo>
		{
			Test("MyApp.Tests.CalculatorTests.Add", "Add"),
			Test("MyApp.Tests.OtherTests.Add", "Add")
		};

		var selected = TestExplorerInterop.FindTestsForTarget(tests, "Add");

		Assert.Equal(2, selected.Count);
	}

	[Fact]
	public void FindTestsForTarget_MatchesTestExplorerId()
	{
		var id = Guid.NewGuid();
		var tests = new List<TestExplorerTestInfo>
		{
			Test("MyApp.Tests.CalculatorTests.Add", "Add", id),
			Test("MyApp.Tests.OtherTests.Add", "Add")
		};

		var selected = TestExplorerInterop.FindTestsForTarget(tests, id.ToString());

		Assert.Equal(id.ToString(), Assert.Single(selected).Id);
	}

	[Fact]
	public void FindTestsForTarget_ReturnsEmpty_WhenNoTestsMatch()
	{
		var selected = TestExplorerInterop.FindTestsForTarget(
			new List<TestExplorerTestInfo> { Test("MyApp.Tests.CalculatorTests.Add", "Add") },
			"MissingTest");

		Assert.Empty(selected);
	}

	private static TestExplorerTestInfo Test(string fullyQualifiedName, string displayName, Guid? id = null) => new()
	{
		Id = (id ?? Guid.NewGuid()).ToString(),
		FullyQualifiedName = fullyQualifiedName,
		DisplayName = displayName
	};
}
