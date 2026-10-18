TOOLTOEMAIL CHO MAC - HUONG DAN NGAN

A. TAO UNG DUNG MAC TU MAY WINDOWS (khong can co may Mac)
1. Tao tai khoan mien phi tai github.com, tao mot repository MOI (de Public thi chay mien phi khong gioi han).
   Luu y: KHONG dua file config.ini co mat khau that len GitHub (dung ban config.ini trong thu muc nay, mat khau de trong).
2. Bam "uploading an existing file" va keo THA TOAN BO noi dung thu muc ToolToEmail_v2
   (app.py, build.bat, build_mac.command, body.txt ... va thu muc .github). Bam Commit changes.
   Neu GitHub khong nhan thu muc .github: Add file > Create new file, dat ten
   ".github/workflows/build-mac.yml" roi dan noi dung file build-mac.yml vao.
3. Vao tab Actions > "Build macOS app" > Run workflow. Doi khoang 3-5 phut den khi co dau tich xanh.
4. Keo xuong muc Artifacts, tai 2 file:
     ToolToEmail-Mac-AppleSilicon  (Mac dung chip M1/M2/M3/M4)
     ToolToEmail-Mac-Intel         (Mac doi cu dung chip Intel)
   Nguoi dung Mac xem chip: Menu Apple > About This Mac.

B. CACH CHAY TREN MAC
1. Giai nen file zip, keo ToolToEmail.app vao thu muc Applications.
2. Lan dau mo, macOS se bao khong xac minh duoc nha phat trien (vi ung dung chua duoc Apple ky).
   - Bam chuot phai (Control + click) vao ToolToEmail > Open > Open.
   - Neu van bi chan (macOS Sequoia tro len): System Settings > Privacy & Security > keo xuong,
     bam "Open Anyway" canh ten ToolToEmail.
   - Hoac mo Terminal va go:  xattr -cr /Applications/ToolToEmail.app
3. Du lieu (config.ini, body.txt, nhat ky send_log.db) nam trong thu muc  ~/ToolToEmail  (thu muc Home).
   Trong ung dung co san nut "Open data folder", "Open config.ini", "Edit email body".
4. Neu dung noi dung email HTML (.html): mo TextEdit > Settings > Open and Save > tich
   "Display HTML files as HTML code instead of formatted text" de sua duoc ma HTML.

C. NEU CO SAN MAY MAC
Mo Terminal trong thu muc nay va chay:  bash build_mac.command  (can Python 3.10+ tu python.org).
