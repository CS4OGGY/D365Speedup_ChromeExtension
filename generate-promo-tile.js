const { execSync } = require("child_process");
const fs = require("fs");
const os = require("os");
const path = require("path");

const dst = "C:\\Users\\anktl\\Pictures\\D365SpeedUp_Store\\";
fs.mkdirSync(dst, { recursive: true });

const iconPath = "C:\\Extensions\\D365Speedup_Google\\assets\\icon128.png";

function makeScript(w, h, outFile) {
  return `
Add-Type -AssemblyName System.Drawing

$bmp = New-Object System.Drawing.Bitmap(${w}, ${h})
$g = [System.Drawing.Graphics]::FromImage($bmp)
$g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
$g.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::AntiAliasGridFit

# Background dark
$bgBrush = New-Object System.Drawing.Drawing2D.LinearGradientBrush(
  (New-Object System.Drawing.Point(0, 0)),
  (New-Object System.Drawing.Point(${w}, ${h})),
  [System.Drawing.Color]::FromArgb(255, 18, 18, 24),
  [System.Drawing.Color]::FromArgb(255, 28, 28, 38)
)
$g.FillRectangle($bgBrush, 0, 0, ${w}, ${h})

# Accent bar top
$accentBrush = New-Object System.Drawing.Drawing2D.LinearGradientBrush(
  (New-Object System.Drawing.Point(0, 0)),
  (New-Object System.Drawing.Point(${w}, 0)),
  [System.Drawing.Color]::FromArgb(255, 80, 140, 255),
  [System.Drawing.Color]::FromArgb(255, 140, 80, 255)
)
$g.FillRectangle($accentBrush, 0, 0, ${w}, 6)

# Icon (centred vertically)
$iconSize = ${Math.round(h * 0.28)}
$iconX = ${Math.round(w * 0.06)}
$iconY = [int]((${ h } - $iconSize) / 2) - 20
$icon = [System.Drawing.Image]::FromFile("${iconPath}")
$g.DrawImage($icon, $iconX, $iconY, $iconSize, $iconSize)
$icon.Dispose()

$textX = $iconX + $iconSize + ${Math.round(w * 0.04)}

# Title
$titleFont = New-Object System.Drawing.Font("Segoe UI", ${Math.round(h * 0.1)}, [System.Drawing.FontStyle]::Bold)
$titleBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(255, 240, 240, 240))
$g.DrawString("D365 SpeedUp", $titleFont, $titleBrush, $textX, ($iconY + 4))

# Subtitle
$subFont = New-Object System.Drawing.Font("Segoe UI", ${Math.round(h * 0.042)}, [System.Drawing.FontStyle]::Regular)
$subBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(255, 100, 160, 255))
$g.DrawString("Chrome Extension  |  Dynamics 365 productivity tools", $subFont, $subBrush, $textX, ($iconY + ${Math.round(h * 0.115)}))

# Tagline
$tagFont = New-Object System.Drawing.Font("Segoe UI", ${Math.round(h * 0.038)}, [System.Drawing.FontStyle]::Regular)
$tagBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(200, 160, 160, 160))
$g.DrawString("Accelerate customization and development with smart tools and generators.", $tagFont, $tagBrush, $textX, ($iconY + ${Math.round(h * 0.2)}))

# Badge
$badgeY = ${h} - ${Math.round(h * 0.2)}
$badgeRect = New-Object System.Drawing.Rectangle($iconX, $badgeY, 140, 28)
$badgeFill = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(40, 80, 140, 255))
$g.FillRectangle($badgeFill, $badgeRect)
$badgePen = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(120, 80, 140, 255), 1)
$g.DrawRectangle($badgePen, $badgeRect)
$badgeFont = New-Object System.Drawing.Font("Segoe UI", ${Math.round(h * 0.036)}, [System.Drawing.FontStyle]::Bold)
$badgeText = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(255, 100, 160, 255))
$g.DrawString("Free  v1.0", $badgeFont, $badgeText, ($iconX + 10), ($badgeY + 5))

# Save JPEG
$codec = [System.Drawing.Imaging.ImageCodecInfo]::GetImageEncoders() | Where-Object { $_.MimeType -eq 'image/jpeg' }
$params = New-Object System.Drawing.Imaging.EncoderParameters(1)
$params.Param[0] = New-Object System.Drawing.Imaging.EncoderParameter([System.Drawing.Imaging.Encoder]::Quality, [long]95)
$bmp.Save("${outFile}", $codec, $params)

$g.Dispose()
$bmp.Dispose()
Write-Host "Saved: ${outFile}"
`;
}

const tiles = [
  { w: 440,  h: 280, name: "promo_small_440x280.jpg" },
  { w: 1400, h: 560, name: "promo_marquee_1400x560.jpg" },
];

const tmpFile = path.join(os.tmpdir(), "gen-promo-tile.ps1");

tiles.forEach(({ w, h, name }) => {
  const outFile = dst + name;
  const psScript = makeScript(w, h, outFile);

  fs.writeFileSync(tmpFile, psScript, "utf8");
  execSync(`powershell -NoProfile -ExecutionPolicy Bypass -File "${tmpFile}"`, { stdio: "inherit" });
});

fs.unlinkSync(tmpFile);
