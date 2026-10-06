---
description: Phan tich video YouTube/local voi skill watch
---

Phan tich video duoc chi dinh: $ARGUMENTS

Quy trinh bat buoc:
1. Load skill `watch` qua skill tool (name: "watch").
2. Chay script de lay transcript + frames (Windows dung `python`, them `--no-install` vi deps da co san ffmpeg/ffprobe/yt-dlp):
`python .opencode/skills/watch/scripts/watch.py "$ARGUMENTS" --no-install --mode balanced`
- Neu user hoi tong quan: them `--mode fast`
- Neu user can chi tiet visual / doc slide / code tren man hinh: `--mode accurate --resolution 768`
- Neu user chi hoi 1 doan: them `--start MM:SS --end MM:SS`
3. Doc transcript truoc (da co trong output). Chi Read nhung frame thuc su can thiet, khong doc het.
4. Tra loi kem timestamp dang [MM:SS].
5. Khong chay setup.sh (chi ho tro macOS/Linux).

Vi du goi:
/watch https://youtu.be/abc123
/watch E:\video\clip.mp4 --start 05:00 --end 06:00
