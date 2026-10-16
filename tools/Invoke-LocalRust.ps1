param(
    [string]$ToolchainRoot = (Join-Path $PSScriptRoot '../../deckcraft-toolchain'),
    [string]$FontsRoot = (Join-Path $PSScriptRoot '../../deckcraft-fonts'),
    [Parameter(ValueFromRemainingArguments = $true)][string[]]$CargoArguments
)
$ErrorActionPreference = 'Stop'
$env:CARGO_HOME = Join-Path ([IO.Path]::GetFullPath($ToolchainRoot)) 'cargo'
$env:RUSTUP_HOME = Join-Path ([IO.Path]::GetFullPath($ToolchainRoot)) 'rustup'
$cargoExecutable = Join-Path $env:CARGO_HOME 'bin/cargo.exe'
if (-not (Test-Path -LiteralPath $cargoExecutable)) { throw 'Prepare a local Rust toolchain or pass -ToolchainRoot.' }
$env:PATH = "$(Join-Path $env:CARGO_HOME 'bin');$env:PATH"
$env:CARGO_BUILD_JOBS = '4'
if (Test-Path -LiteralPath $FontsRoot) {
    $env:CRAFT_FONTS_DIR = [IO.Path]::GetFullPath($FontsRoot)
    $env:CRAFT_FONTS_REQUIRED = '1'
}
Push-Location (Join-Path $PSScriptRoot '..')
try {
    & $cargoExecutable @CargoArguments
    exit $LASTEXITCODE
} finally {
    Pop-Location
}
