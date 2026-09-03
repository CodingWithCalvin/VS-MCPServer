using CodingWithCalvin.MCPServer.Services;
using Xunit;

namespace CodingWithCalvin.MCPServer.Tests;

public class BuildConfigurationTests
{
    [Theory]
    [InlineData("Debug", "Any CPU", "debug", "any cpu", true)]
    [InlineData("Release", "x64", "Debug", "x64", false)]
    [InlineData("Release", "x64", "Release", "ARM64", false)]
    public void MatchesBuildConfiguration_RequiresExactConfigurationAndPlatform(
        string candidateConfiguration,
        string candidatePlatform,
        string configuration,
        string platform,
        bool expected)
    {
        Assert.Equal(
            expected,
            VisualStudioService.MatchesBuildConfiguration(
                candidateConfiguration,
                candidatePlatform,
                configuration,
                platform));
    }
}
