# ⭐ Star Wars Excel Maze

*A long time ago, in an excel spreadsheet far, far away…*

Used Microsoft Excel and VBA to make an old-school maze for a course. Travel with the 'Millennium Falcon' through the cells in the maze while collecting hyperfuel and dodging Death Stars.


## How to play

1. Download [`StarWarsMaze.xlsm`](StarWarsMaze.xlsm) and open it in **desktop**
3. Click **Start Game**.
4. Navigate using **Up / Down / Left / Right** buttons.

### Rules

| Square | What it means |
|---|---|
| 🟨 Yellow | Safe routes through galaxy |
| 💥 Light blue | Hyperfuel: +1 |
| 🌑 Red | Death Star: press **LightSpeed** to avoid detection (costs 3 fuel). Without enough fuel you're caught and sent back to the Rebel base |
| 👽 Dark Blue | Dagobah: Master Moda's advice. You need it to finish |
| 👸 Magenta | Princess Meia: you have reached the end of the game |
| ⬛ Black / white | Deep space. Cannot enter. |

## How it works

Logic/code of the game exported to [`src/`](src):

- [`Movement.bas`](src/Movement.bas): `StartGame_Click` places the characters and fuel; `Up/Down/Left/Right_Click` move the player and check the colour of the target cell to decide what happens; `Reset_Click` restores the board from a hidden copy sheet.
- [`LightSpeed.bas`](src/LightSpeed.bas): checks each direction for an adjacent Death Star and either jumps past it (spending 3 fuel) or sends you back to the start.

The game state lives in the sheet itself: the player's position is the selected cell, cell colours define the map, and `AF7` / `AF9` hold the fuel count and whether you've received advice.

## Built with

Microsoft Excel · VBA
