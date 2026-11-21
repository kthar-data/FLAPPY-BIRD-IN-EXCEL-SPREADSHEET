# Hello, and welcome to Flappy Bird in Excel!!!

A playable Flappy Bird-style game that runs from a single VBA module inside desktop Excel. It opens in its own small window and draws with Windows GDI, so it stays smooth instead of fighting Excel's cell repainting.

## Features :)

- Smooth 60 fps rendering in its own window (no cell or shape repainting)
- Fixed-timestep physics, so the game feels the same on fast and slow machines
- Tap-proof input: key presses are read from the message queue, so short taps are never missed
- Score, session best, game-over screen and instant restart
- One plain-text file you can read before you import it

## Requirements

- Windows desktop Excel 2010 or newer (32-bit or 64-bit)
- Not supported: Excel for Mac, Excel for the web, LibreOffice, Google Sheets

## Install (Easy)

1. Open a new workbook in Excel.
2. Press `Alt+F11` to open the VBA editor.
3. Choose **File > Import File...** and select `FlappyBird.bas`.
4. Close the editor, press `Alt+F8`, select `SetupFlappyBird` and click **Run**. This builds a launcher sheet with a **PLAY** button.
5. Click **PLAY**.
6. To keep the game, save the workbook as **Excel Macro-Enabled Workbook (`.xlsm`)**.

You can also skip the launcher sheet and run `PlayFlappyBird` directly from `Alt+F8`.

## Controls (simple)

| Input | Action |
| --- | --- |
| `Space`, `Up arrow` or mouse click | Flap (and start / restart) |
| `Esc` or the window's close button | Quit |

## Tuning

Constants near the top of `FlappyBird.bas` control the feel and difficulty:

| Constant | Default | Effect |
| --- | --- | --- |
| `PIPE_GAP` | 165 | Height of the opening (smaller = harder) |
| `PIPE_SPEED` | 140 | How fast pipes scroll, in pixels per second |
| `PIPE_SPACING` | 225 | Distance between pipes |
| `MAX_SHIFT` | 100 | Largest vertical jump between two gaps |
| `GRAV` / `FLAP_V` | 1500 / -430 | Gravity and flap strength |
| `WIN_SCALE` | 1 | Set to 2 for a double-size window on high-resolution screens |

## Troubleshooting

- **"Ambiguous name detected" or duplicate-name error on import:** an older `FlappyBird` module is still in the project. Right-click it in the VBA editor, choose **Remove**, click **No** when asked to export, then import again.
- **Macros are disabled:** save the workbook as `.xlsm` and enable macros when Excel asks, or put the file in a Trusted Location.
- **Window is tiny on a 4K screen:** set `WIN_SCALE` to `2`.
- **Compile error on a specific line:** open an issue with the exact error text and the highlighted line.

## How it works

The game never touches the worksheet while you play. Each frame is drawn into an off-screen bitmap with GDI and copied to a plain Windows window in one step. Keyboard and mouse input comes straight from the message queue, and physics runs on a fixed 120 Hz timestep decoupled from the 60 fps display loop.

## Safety

The module is plain text. It makes no network connections, reads and writes no files, runs no shell commands and touches no registry or clipboard. Its only external calls are 34 Windows functions for creating the window, drawing, timing and reading the game's own key presses. The only worksheet it changes is the launcher sheet named `Flappy Bird`, which `SetupFlappyBird` clears and rebuilds, so do not rename one of your own sheets to that name.

As with any macro-enabled file, only enable macros in workbooks from sources you trust.

## License

MIT. See [LICENSE](LICENSE).

This is an independent fan-made project and is not affiliated with or endorsed by the creator or publisher of the original Flappy Bird.
