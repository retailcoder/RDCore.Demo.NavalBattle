# RDCore.Demo.NavalBattle
<sup>_Ce document est disponible en [français](./README.fr.md)_</sup>

![VIVAT CUCUMIS](./assets/vivat-cucumis-stonecore.png)

RD-VBA Naval Battle is a naval battle game written in VBA by the author of Rubberduck to show object-oriented VBA: interfaces and `Implements`, factory methods, events between classes, and custom errors. It was legacy Rubberduck's showcase.

This repository holds the game's port to the [RDCore](https://github.com/rubberduck-vba/RDCore) platform, which opens, runs and debugs VBA without Office hosting it.

## Project status

RD-VBA Naval Battle is in **pre-alpha** development. It plays in a window on Windows, and in the console elsewhere: a person against the computer, or the computer against itself.

## Running the game

The game runs with `rdc`, the RDCore command line, built from the [RDCore](https://github.com/rubberduck-vba/RDCore) repository. From the folder of this repository:

```
rdc run .
```

On Windows, the game opens a window with both grids. Pick who plays each grid, Human or one of the computer's players, then click **Ready**. A person places their five ships on their own grid: point at a cell to move the ship there, right-click to turn it, and click to place it. Click **Ready** again, then fire at the other grid by clicking its cells. When the computer plays both grids, the whole game plays out in the window. Closing the window ends the game. With `/console` after `--` (see below), the game plays in the console instead.

In the console, the game asks who plays each grid: Human, Random, Fair-play or Merciless (the first letter is enough). A person places their five ships, then takes turns with the computer, one shot each, until a fleet is sunk. A cell is a column letter and a row number, like `B3`; a ship goes across the grid from its first cell, or down with a `V`: `B3 V`. An empty answer ends the game. On Windows, the questions are dialog boxes.

After `--`, `/player1 <player>` and `/player2 <player>` name the players (the window opens with them already picked), `/console` plays in the console on Windows too, and `/seed <number>` fixes the random seed, so that a game can be played again move for move:

```
rdc run . -- /player1 merciless /player2 fairplay /seed 42
```

## Licence

RD-VBA Naval Battle is licensed under the **MIT** licence: see [LICENSE](LICENSE) and [NOTICE.md](NOTICE.md).

## Contributions

- Individuals: open, under the [RDCore CLA](https://github.com/rubberduck-vba/RDCore/blob/main/CLA.md). One signature covers every RDCore repository.
- Corporate: planned.

## See also

- [RDCore](https://github.com/rubberduck-vba/RDCore): the platform, its runtime and its language server

<hr/>
<h6 align='center'>V I V A T ❤️ C U C U M I S ™</h6>
<p align='center' style='font-size:8pt;'>
<small>© Copyright <strong>9562-7303 Québec inc.</strong> (2026)<br/><em>Seul, &quot;Rubberduck&quot; est utilisé pour fins de référence au projet open-source legacy <strong>utilisé publiquement ainsi depuis 2015</strong> et sans lien ni affiliation avec tout tiers détenteur d'une marque semblable dans quelque juridiction que ce soit.<br/>&quot;Rubberduck VBA&quot;, &quot;RDCore&quot; et &quot;VIVAT CUCUMIS&quot; sont des marques de commerce revendiquées par 9562-7303 Québec inc. (en attente); Toutes les marques appartiennent à leur détenteur respectif.<br/>RDCore.Demo.NavalBattle n'est pas un produit de Microsoft et n'est pas affilié à Microsoft, ni directement, ni indirectement.<br/>BATTLESHIP est une marque de commerce de Hasbro, Inc. RDCore.Demo.NavalBattle n'est ni affilié à Hasbro, Inc., ni commandité ni approuvé par celle-ci.</em><br/><br/><em>If used alone, &quot;Rubberduck&quot; is used as a reference to the legacy open-source project <strong>the same way it has been used publicly since 2015</strong> and without any links or affiliation with any third-party trademark holders of a similar trademark in any jurisdiction.<br/>&quot;Rubberduck VBA&quot;, &quot;RDCore&quot; and &quot;VIVAT CUCUMIS&quot; are trademarks claimed by 9562-7303 Québec inc. (pending). All trademarks belong to their respective owners.<br/>RDCore.Demo.NavalBattle is not a Microsoft product and is not affiliated with Microsoft, directly or indirectly.<br/>BATTLESHIP is a trademark of Hasbro, Inc. RDCore.Demo.NavalBattle is not affiliated with, sponsored by or endorsed by Hasbro, Inc.</em></small>
</p>
