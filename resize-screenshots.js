const { execSync } = require("child_process");
const fs = require("fs");

const src = "C:\\Users\\anktl\\Pictures\\Screenshots\\";
const dst = "C:\\Users\\anktl\\Pictures\\D365SpeedUp_Store\\";

fs.mkdirSync(dst, { recursive: true });

const screenshots = [
  "Screenshot 2026-03-13 172829.png",
  "Screenshot 2026-03-13 172858.png",
  "Screenshot 2026-03-13 172918.png",
  "Screenshot 2026-03-13 172935.png",
  "Screenshot 2026-03-13 173236.png"
];

screenshots.forEach((f, i) => {
  const inp = src + f;
  const out = dst + "screenshot_" + (i + 1) + ".jpg";
  const ps = [
    "Add-Type -AssemblyName System.Drawing;",
    `$img = [System.Drawing.Image]::FromFile('${inp}');`,
    "$bmp = New-Object System.Drawing.Bitmap(1280,800);",
    "$g = [System.Drawing.Graphics]::FromImage($bmp);",
    "$g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic;",
    "$g.DrawImage($img,0,0,1280,800);",
    "$codec = [System.Drawing.Imaging.ImageCodecInfo]::GetImageEncoders() | Where-Object { $_.MimeType -eq 'image/jpeg' };",
    "$params = New-Object System.Drawing.Imaging.EncoderParameters(1);",
    "$params.Param[0] = New-Object System.Drawing.Imaging.EncoderParameter([System.Drawing.Imaging.Encoder]::Quality, 92L);",
    `$bmp.Save('${out}', $codec, $params);`,
    "$g.Dispose(); $bmp.Dispose(); $img.Dispose();"
  ].join(" ");

  execSync(`powershell -NoProfile -Command "${ps}"`);
  console.log("Saved: " + out);
});

console.log("\nAll done! Files saved to: " + dst);
