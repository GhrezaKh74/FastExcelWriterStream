#!/bin/bash
# Build and pack FastExcelWriter NuGet package

set -e

echo "=== Building FastExcelWriter ==="
dotnet build src/FastExcelWriter/FastExcelWriter.csproj -c Release

echo ""
echo "=== Running Tests ==="
dotnet test tests/FastExcelWriter.Tests/FastExcelWriter.Tests.csproj -c Release --no-build

echo ""
echo "=== Packing NuGet ==="
dotnet pack src/FastExcelWriter/FastExcelWriter.csproj -c Release -o ./nupkg

echo ""
echo "✅ Done! Package is in ./nupkg/"
echo "   To install locally:  dotnet add package FastExcelWriter --source ./nupkg"
echo "   To publish to NuGet: dotnet nuget push ./nupkg/*.nupkg --api-key YOUR_KEY --source https://api.nuget.org/v3/index.json"
