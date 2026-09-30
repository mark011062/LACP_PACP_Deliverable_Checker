FFPROBE VIDEO VALIDATION

This version adds an FFprobe integrity pass over video files in the selected Video folder.

For a self-contained Windows EXE:
1. Obtain the official FFmpeg Windows build you use in your organization.
2. Copy ffprobe.exe into this Acme_Checker folder (beside BUILD.bat).
3. Double-click BUILD.bat.
4. The .spec file bundles ffprobe.exe into the generated EXE.

If ffprobe.exe is not bundled, the app will also look for ffprobe on the computer's PATH.
If FFprobe cannot be found, the normal database/media existence checks still run and the report clearly states that integrity validation was skipped.

Integrity validation flags:
- files FFprobe cannot read
- files with no video stream
- files with missing/invalid duration
- zero-second videos
