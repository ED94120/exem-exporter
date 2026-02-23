#**Extraction capteurs EXEM**

##**Objectif**

Ce script a une finalité exclusivement informative, éducative et de recherche.
Il permet d’analyser l’évolution temporelle des niveaux d’exposition mesurés par les capteurs EXEM affichés sur le site ANFR de l’Observatoire des ondes (https://www.observatoiredesondes.com/fr/).
L’objectif est d’étudier le comportement des niveaux d’exposition en fonction du temps et de l’environnement du capteur.

Le script respecte les modalités d’utilisation du site :
- aucune requête automatisée vers les serveurs n’est effectuée,
- aucune base de données ni API n’est interrogée,
- seules les informations déjà visibles à l’écran sont lues localement dans le navigateur.

##**Conditions de fonctionnement**

Le script fonctionne uniquement si :
- les points de mesure sont visibles sur le graphique temps–mesure,
- la courbe affichée comporte des marqueurs individuels,
- la période affichée est suffisamment courte pour que les points soient activés (en général 7 jours maximum).

Si la courbe est affichée en mode lissé sans points visibles, l’extraction est impossible.

##**Principe technique**

Le programme :
- détecte les points de mesure visibles dans le graphique,
- simule le survol de chaque point,
- lit le contenu des pop-ups,
- extrait la date, l’heure et la valeur en V/m,
- trie les données chronologiquement,
- calcule les statistiques (min, moyenne, max),
- génère un fichier CSV structuré.

Le script ne lit aucune donnée cachée et n’accède pas aux serveurs du site.
Il se contente d’automatiser la lecture des informations déjà affichées à l’écran.
