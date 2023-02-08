![Project Preview Image](/media/preview.png)

# Microsoft Excel Blackjack _(Twenty-One)_

The purpose of this project was to test the built-in capabilities of Microsoft Excel tool _(**without using VBA**)_ to perform non-routine tasks. A simplified, single-player version of the popular blackjack game was developed for this purpose.

## 🤔 How to play Blackjack

### Object of the Game

The goal of the participant is to get as close to 21 as possible, without going over.

### Card values

Whether an ace is worth 1 or 11 depends on the player. Face cards are worth 10 and any other card is worth its pip value.

### The play

The player goes first and decides whether to "stand" (not ask for another card) or "hit" (ask for another card in order to reach 21). Hence, a player may stand on their two cards or ask the dealer for additional cards, one at a time, until they either decide to stand on the total (if it is under 21) or go "bust" (if it is over 21).

## 🖨️ Implementation

[Manually recalculating](https://support.microsoft.com/en-us/office/change-formula-recalculation-iteration-or-precision-in-excel-73fc7dac-91cf-4d36-86e8-67124f6bcce4) the cells in the sheets is the main option used in the implementation of the game. This option prevents circular references between functions and performs calculations only when the user manually presses the `F9` key.

There are 5 cards on each player's board, with the player starting with two cards and the dealer with one. A player's card total and corresponding win/loss message will be displayed when the `F9` key is pressed.

### Three sheets are in the spreadsheet

#### Calculations

Where all the "behind the scenes" calculations are made:

- Shuffling th order of the cards in the pack.
- Distribution of the cards to the dealer and the participant.
- In each draw, the sum of the cards is calculated for each player.

The main functions of the game are:

1. `Changed Ace?` which calculates whether the sum of aces should change from 11 to 1 to allow the game to continue.
2. `Sum` which calculates the player's sum and whether he reached the sum of 21 and won, or exceeded it and lost (A tie between the dealer and the player is also considered a loss).

#### How to play

Provides a brief explanation of the game.

#### Blackjack

Where the game itself takes place, serves as the user interface for the game.

## 🕹️ Getting Started

### Dependencies

- Microsoft Office Excel 2003, or later.

### Start a new game

1. Launch the Microsoft Excel spreadsheet on the `Blackjack` sheet.
2. Set the "Restart" label to `TRUE` and the "Draw more" label to `HIT`.
3. To draw more cards, press `F9`.
4. When you are done asking for more cards, change "Draw more" label to `STAND`.
5. To find out who won, press `F9` one last time.

## 📝 License

Copyright © 2019, E-RELevant
All rights reserved.

This source code is licensed under the GPL-v3.0 license found in the [LICENSE](LICENSE) file in the root directory of this source tree.
