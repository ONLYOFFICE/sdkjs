# Desktop media command regression tests

Run from the sdkjs root with Node.js 20 or newer:

```sh
node tests/desktop/media-player.test.cjs
```

These dependency-free tests execute the real `common/apiBase.js`. The animation
integration case also loads `HistoryCommon.js` and `Timing.js` and activates a
real `CCmd`, as happens after a Next event starts a PowerPoint media effect.
Editor geometry is stubbed; the native bridge is replaced by a small player
model with the initialization and selection requirements of
`desktop-sdk/ChromiumBasedEditors/lib/qt_wrapper/src/qcefview_media.cpp`.

Coverage includes translating `playFrom(0.0)`, initializing an unselected audio
shape, repeating playback, switching sources, existing control commands,
geometry update deduplication, unavailable geometry, and a browser without the
desktop bridge. Ordinary play retains the existing player. Nonzero `playFrom`
offsets remain unsupported by the native bridge and are not translated to a
restart at zero. These command tests do not exercise native playback.

## Desktop output regression (Linux)

With Node.js 22 or newer, `xvfb-run`, `pactl`, `parec` (pulseaudio-utils), a running
PulseAudio-compatible server, and an assembled Desktop Editors build containing
this SDK:

```sh
node tests/desktop/audio-playback.mjs /path/to/DesktopEditors
```

The test opens the included `fixtures/next-click-audio.pptx` in an isolated profile
and display, starts a slideshow, and invokes its Next action without selecting
the media picture. It captures only its own temporary null sink's monitor.
It checks silence before Next, the decoded 440 Hz tone on slide 1, the 880 Hz tone
on slide 2, replay of slide 1 after restarting the slideshow, and slideshow exit.
The test removes its sink and closes its editor. Port 8080 must be free for the
editor's debugging endpoint. Diagnostic logs remain in the reported temporary
profile directory. `PACTL` and `PAREC` can override the utility paths.

The PPTX contains generated sine tones and a generated icon, with a
`p:cmd type="call" cmd="playFrom(0.0)"` click effect on each slide. It contains no
user presentation content. To reproduce manually, open it, start the slideshow,
and press Next once to hear the tone; the slide should remain displayed.
To regenerate the fixture, install `python-pptx` and `ffmpeg`, then run:

```sh
python3 tests/desktop/generate-audio-fixture.py
```

Output capture verifies PPTX import, animation triggering, native player creation,
MP3 decoding and output routing. It does not test physical speakers, other
operating systems, nonzero seeking, or simultaneous media playback.
