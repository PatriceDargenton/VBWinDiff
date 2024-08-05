# VBWinDiff
Interface d'options pour le comparateur WinDiff (et WinMerge)

VBWinDiff permet d'ajouter des options utiles pour le comparateur de fichier texte WinDiff, ainsi que WinMerge : l'idée, c'est d'effectuer un prétraitement des fichiers textes à comparer, de façon à ignorer des détails qui surchargent la comparaison via WinDiff ou WinMerge. Exemple : si vous avez des espaces insécables dans un fichier et pas dans l'autre, WinDiff vous affiche des tas de différences qui ne vous intéressent sans doute pas. De même pour la casse (majuscule/minuscule), les accents, la ponctuation. Il reste cependant des limitations : si les phrases sont trop longues, ou bien s'il y a un retour à la ligne dans une phrase, la détection ne fonctionne plus. Du coup, il y a une option permettant la comparaison mot à mot : WinDiff retrouve alors l'ensemble des différences sans ternir compte de la ponctuation, seulement des mots, ce qui est utile par exemple pour comparer deux versions d'un texte.

## Table des matières
- [Utilisation](#utilisation)
- [Limitations](#limitations)
- [Projets](#projets)
- [Versions](#versions)
- [Liens](#liens)

## Utilisation
Placer un raccourci vers VBWinDiff.exe dans le dossier SendTo de Windows (droit admin. requis), et envoyer deux fichiers à comparer vers VBWinDiff via l'explorateur de fichier de Windows.

Chemin du dossier SendTo depuis Windows XP :
"C:\Documents and Settings\[MonCompte]\SendTo"

Chemin du dossier SendTo pour Windows Vista :
"C:\Documents and Settings\[MonCompte]\Application Data\Microsoft\Windows\SendTo"

Chemin du dossier SendTo depuis Windows 7 et supérieurs :
"C:\Users\[MonCompte]\AppData\Roaming\Microsoft\Windows\SendTo"

Note : si vous lancez VBWinDiff sans argument, VBWinDiff se chargera de gérer ce raccourci : installation / désinstallation.

Ratio de pagination : c'est une option pour paginer en considérant que les textes ont un contenu identique, même s'ils ne sont pas de la même longueur : soit la différence de taille est due par exemple à la présence de sauts de lignes à chaque ligne dans l'un des deux fichiers (dans ce cas appliquer le ratio), soit l'un des textes a du contenu en plus ; dans ce cas, soit se contenu est à la fin, soit on ne sait pas, et alors dans ce cas il peut être utile aussi d'appliquer un ratio.

Comparaison mot à mot : Si vous devez comparer deux textes dont l'un a des retours à la ligne intempestif, il faut alors tenter une comparaison mot à mot (sans la ponctuation, ni les phrases, juste les mots successifs au "kilomètre"). Si en plus la césure des mots a été appliquée, VBWinDiff peut alors corriger automatiquement ces mots coupés sur une base statistique : si le mot recollé est plus fréquent que la somme des tronçons, alors c'est une césure probable que l'on peut donc corriger. Dans ce mode, la taille des fichiers après traitement est affichée aussi, pour donner une idée de la similitude des deux documents. Pour activer cette option, décocher l'option Paragraphe.

## Limitations
- WinDiff ne gère pas l'unicode (par exemple les caractères accentués son difficiles à lire) ;
- En mode mot à mot, WinDiff ne compare rapidement que des fichiers de moins de 100 Ko (sinon vous pouvez attendre longtemps, longtemps...) ;
- WinDiff ne gère pas les sauts de lignes dans une phrase (par exemple si l'un des fichiers provient d'un copié/collé via un fichier pdf) : les phrases ne sont plus détectées comme identiques, on est obligé de découper tous les mots pour que WinDiff puisse retrouver les différences.

## Projets

## Versions

Voir le [Changelog.md](Changelog.md)

## Liens

- [WinDiff](https://en.wikipedia.org/wiki/WinDiff) : [version 5.2.3790.0 du 24/03/2003](http://www.grigsoft.com/windiff.zip) livrée avec Windows 2003 (Microsoft Source Code Samples).

Documentation d'origine complète : [VBWinDiff.html](http://patrice.dargenton.free.fr/CodesSources/VBWinDiff.html)