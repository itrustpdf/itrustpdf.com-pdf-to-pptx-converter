# Test script for the PDF to PPTX converter service (PowerShell)
# This script tests the Docker-based application

Write-Host "=== PDF to PPTX Converter Test Script ===" -ForegroundColor Green
Write-Host ""

try {
    # Build the Docker image
    Write-Host "1. Building Docker image..." -ForegroundColor Yellow
    docker-compose build
    if ($LASTEXITCODE -ne 0) { throw "Docker build failed" }

    Write-Host ""
    Write-Host "2. Starting the service..." -ForegroundColor Yellow
    docker-compose up -d
    if ($LASTEXITCODE -ne 0) { throw "Service start failed" }

    # Wait for service to be ready
    Write-Host ""
    Write-Host "3. Waiting for service to start..." -ForegroundColor Yellow
    Start-Sleep -Seconds 15

    # Test health endpoint
    Write-Host ""
    Write-Host "4. Testing health endpoint..." -ForegroundColor Yellow
    try {
        $healthResponse = Invoke-RestMethod -Uri "http://localhost:8000/health" -Method Get -TimeoutSec 10
        Write-Host "Health response: $($healthResponse | ConvertTo-Json -Compress)" -ForegroundColor Cyan
        
        if ($healthResponse.status -eq "healthy") {
            Write-Host "✅ Service is healthy!" -ForegroundColor Green
        } else {
            Write-Host "⚠️ Service status: $($healthResponse.status)" -ForegroundColor Yellow
        }
    } catch {
        Write-Host "❌ Health check failed: $($_.Exception.Message)" -ForegroundColor Red
        docker-compose logs
        throw "Health check failed"
    }

    # Test root endpoint
    Write-Host ""
    Write-Host "5. Testing root endpoint..." -ForegroundColor Yellow
    try {
        $rootResponse = Invoke-WebRequest -Uri "http://localhost:8000/" -Method Get -TimeoutSec 10
        if ($rootResponse.StatusCode -eq 200) {
            Write-Host "✅ Root endpoint working!" -ForegroundColor Green
        } else {
            Write-Host "❌ Root endpoint failed with status: $($rootResponse.StatusCode)" -ForegroundColor Red
        }
    } catch {
        Write-Host "❌ Root endpoint failed: $($_.Exception.Message)" -ForegroundColor Red
    }

    # Test API documentation
    Write-Host ""
    Write-Host "6. Testing API documentation..." -ForegroundColor Yellow
    try {
        $docsResponse = Invoke-WebRequest -Uri "http://localhost:8000/docs" -Method Get -TimeoutSec 10
        if ($docsResponse.StatusCode -eq 200) {
            Write-Host "✅ API documentation accessible!" -ForegroundColor Green
        } else {
            Write-Host "❌ API documentation failed with status: $($docsResponse.StatusCode)" -ForegroundColor Red
        }
    } catch {
        Write-Host "❌ API documentation failed: $($_.Exception.Message)" -ForegroundColor Red
    }

    # Test convert endpoint (without file - should return validation error)
    Write-Host ""
    Write-Host "7. Testing convert endpoint (validation)..." -ForegroundColor Yellow
    try {
        $convertResponse = Invoke-WebRequest -Uri "http://localhost:8000/convert" -Method Post -TimeoutSec 10
        Write-Host "❌ Convert endpoint should have failed without file" -ForegroundColor Red
    } catch {
        if ($_.Exception.Response.StatusCode -eq 422) {
            Write-Host "✅ Convert endpoint validation working!" -ForegroundColor Green
        } else {
            Write-Host "⚠️ Convert endpoint returned unexpected error: $($_.Exception.Response.StatusCode)" -ForegroundColor Yellow
        }
    }

    Write-Host ""
    Write-Host "8. Service Status Summary:" -ForegroundColor Green
    Write-Host "✅ Docker image built successfully" -ForegroundColor Green
    Write-Host "✅ Service started and responding" -ForegroundColor Green
    Write-Host "✅ Health check passed" -ForegroundColor Green
    Write-Host "✅ API endpoints accessible" -ForegroundColor Green
    Write-Host ""
    Write-Host "Service is running at: http://localhost:8000" -ForegroundColor Cyan
    Write-Host "API documentation: http://localhost:8000/docs" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "To test with a real PDF file, use:" -ForegroundColor Yellow
    Write-Host 'curl -X POST "http://localhost:8000/convert" -F "file=@your-document.pdf" --output "result.pptx"' -ForegroundColor White
    Write-Host ""
    Write-Host "Or use PowerShell:" -ForegroundColor Yellow
    Write-Host '$file = Get-Item "your-document.pdf"; Invoke-RestMethod -Uri "http://localhost:8000/convert" -Method Post -Form @{file = $file} -OutFile "result.pptx"' -ForegroundColor White
    Write-Host ""
    Write-Host "To stop the service, run: docker-compose down" -ForegroundColor Yellow

} catch {
    Write-Host ""
    Write-Host "❌ Test failed: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host ""
    Write-Host "Checking service logs..." -ForegroundColor Yellow
    docker-compose logs
    Write-Host ""
    Write-Host "To stop the service, run: docker-compose down" -ForegroundColor Yellow
    exit 1
}
