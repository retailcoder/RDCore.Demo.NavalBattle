# RDCore.Demo.NavalBattle
<sup>_This document is available in [English](./README.md)_</sup>

![VIVAT CUCUMIS](./assets/vivat-cucumis-stonecore.png)

Bataille navale RD-VBA est un jeu de bataille navale écrit en VBA par l'auteur de Rubberduck pour montrer la programmation orientée objet en VBA : les interfaces et `Implements`, les méthodes de fabrique, les événements entre classes et les erreurs personnalisées. C'était la vitrine de Rubberduck legacy.

Ce référentiel contient le portage du jeu vers la plateforme [RDCore](https://github.com/rubberduck-vba/RDCore), qui ouvre, exécute et débogue du VBA sans qu'Office l'héberge.

## État du projet

Bataille navale RD-VBA est en développement **pré-alpha**. Le jeu se joue dans une fenêtre sous Windows, et dans la console ailleurs : une personne contre l'ordinateur, ou l'ordinateur contre lui-même.

## Lancer le jeu

Le jeu s'exécute avec `rdc`, la ligne de commande de RDCore, compilée à partir du référentiel [RDCore](https://github.com/rubberduck-vba/RDCore). À partir du dossier de ce référentiel :

```
rdc run .
```

Sous Windows, le jeu ouvre une fenêtre avec les deux grilles. Choisissez qui joue chaque grille, Human ou l'un des joueurs de l'ordinateur, puis cliquez sur **Ready**. Une personne place ses cinq navires sur sa propre grille : pointez une case pour y déplacer le navire, cliquez avec le bouton droit pour le tourner, et cliquez pour le placer. Cliquez de nouveau sur **Ready**, puis tirez sur l'autre grille en cliquant sur ses cases. Quand l'ordinateur joue les deux grilles, toute la partie se déroule dans la fenêtre. Fermer la fenêtre termine la partie. Avec `/console` après `--` (voir plus bas), le jeu se joue plutôt dans la console.

Dans la console, le jeu demande qui joue chaque grille : Human, Random, Fair-play ou Merciless (la première lettre suffit). Une personne place ses cinq navires, puis joue à tour de rôle avec l'ordinateur, un tir chacun, jusqu'à ce qu'une flotte soit coulée. Une case est une lettre de colonne et un numéro de ligne, comme `B3`; un navire part de sa première case vers la droite, ou vers le bas avec un `V` : `B3 V`. Une réponse vide termine la partie. Sous Windows, les questions sont des boîtes de dialogue. Le jeu lui-même est en anglais.

Après `--`, `/player1 <joueur>` et `/player2 <joueur>` nomment les joueurs (la fenêtre s'ouvre avec ces joueurs déjà choisis), `/console` fait jouer dans la console sous Windows aussi, et `/seed <nombre>` fixe la graine aléatoire, pour qu'une partie puisse être rejouée coup pour coup :

```
rdc run . -- /player1 merciless /player2 fairplay /seed 42
```

## Licence

Bataille navale RD-VBA est distribuée sous la licence **MIT** : voir [LICENSE](LICENSE) et [NOTICE.md](NOTICE.md).

## Contributions

- Individus : ouvertes, sous le [CLA de RDCore](https://github.com/rubberduck-vba/RDCore/blob/main/CLA.md). Une seule signature couvre tous les référentiels de RDCore.
- Entreprises : prévues.

## Voir aussi

- [RDCore](https://github.com/rubberduck-vba/RDCore) : la plateforme, son _runtime_ et son serveur de langage

<hr/>
<h6 align='center'>V I V A T ❤️ C U C U M I S ™</h6>
<p align='center' style='font-size:8pt;'>
<small>© Copyright <strong>9562-7303 Québec inc.</strong> (2026)<br/><em>Seul, &quot;Rubberduck&quot; est utilisé pour fins de référence au projet open-source legacy <strong>utilisé publiquement ainsi depuis 2015</strong> et sans lien ni affiliation avec tout tiers détenteur d'une marque semblable dans quelque juridiction que ce soit.<br/>&quot;Rubberduck VBA&quot;, &quot;RDCore&quot; et &quot;VIVAT CUCUMIS&quot; sont des marques de commerce revendiquées par 9562-7303 Québec inc. (en attente); Toutes les marques appartiennent à leur détenteur respectif.<br/>RDCore.Demo.NavalBattle n'est pas un produit de Microsoft et n'est pas affilié à Microsoft, ni directement, ni indirectement.<br/>BATTLESHIP est une marque de commerce de Hasbro, Inc. RDCore.Demo.NavalBattle n'est ni affilié à Hasbro, Inc., ni commandité ni approuvé par celle-ci.</em><br/><br/><em>If used alone, &quot;Rubberduck&quot; is used as a reference to the legacy open-source project <strong>the same way it has been used publicly since 2015</strong> and without any links or affiliation with any third-party trademark holders of a similar trademark in any jurisdiction.<br/>&quot;Rubberduck VBA&quot;, &quot;RDCore&quot; and &quot;VIVAT CUCUMIS&quot; are trademarks claimed by 9562-7303 Québec inc. (pending). All trademarks belong to their respective owners.<br/>RDCore.Demo.NavalBattle is not a Microsoft product and is not affiliated with Microsoft, directly or indirectly.<br/>BATTLESHIP is a trademark of Hasbro, Inc. RDCore.Demo.NavalBattle is not affiliated with, sponsored by or endorsed by Hasbro, Inc.</em></small>
</p>
