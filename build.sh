#!/bin/bash
# Build and pack FastExcelWriterStream NuGet package

set -e

echo "=== Building FastExcelWriterStream ==="
dotnet build src/FastExcelWriterStream/FastExcelWriterStream.csproj -c Release

echo ""
echo "=== Running Tests ==="
dotnet test tests/FastExcelWriterStream.Tests/FastExcelWriterStream.Tests.csproj -c Release --no-build

echo ""
echo "=== Packing NuGet ==="
dotnet pack src/FastExcelWriterStream/FastExcelWriterStream.csproj -c Release -o ./nupkg

echo ""
echo "✅ Done! Package is in ./nupkg/"
echo "   To install locally:  dotnet add package FastExcelWriterStream --source ./nupkg"
echo "   To publish to NuGet: dotnet nuget push ./nupkg/*.nupkg --api-key YOUR_KEY --source https://api.nuget.org/v3/index.json"
