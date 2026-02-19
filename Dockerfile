FROM mcr.microsoft.com/dotnet/sdk:8.0-alpine AS build
WORKDIR /app

# Copy project file and restore first (better layer caching)
COPY DocxReview.csproj .
RUN dotnet restore

# Copy source and publish
COPY src/ src/
COPY templates/ templates/
RUN dotnet publish -c Release -o /app/publish --no-restore \
    /p:PublishTrimmed=false \
    /p:PublishSingleFile=false

# Runtime-only Alpine image
FROM mcr.microsoft.com/dotnet/runtime:8.0-alpine
WORKDIR /app
COPY --from=build /app/publish .

ENTRYPOINT ["dotnet", "/app/docx-review.dll"]
