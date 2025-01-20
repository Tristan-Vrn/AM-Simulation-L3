# Gestion de Portefeuille d'Investissement

## Description du Projet

Ce projet contient un **outil de gestion de portefeuille d’investissement** destiné à un client ayant confié une somme initiale de **1 000 000 € au 1er janvier 2005**. L’outil permet de suivre et d’analyser la composition, la valeur, la performance et les mesures de risque d’un portefeuille composé d’actions et d’obligations.  

### Fonctionnalités Principales
L’outil offre les fonctionnalités suivantes :
1. **Suivi du portefeuille :**
   - Composition par poche d’investissement (actions et obligations).
   - Valeur de chaque poche à une date donnée.
   - Performance depuis le début de la collaboration et depuis la dernière requête.

2. **Analyse des risques :**
   - Volatilité des actions individuelles et du portefeuille global.
   - Duration et duration modifiée des obligations.

3. **Synthèse globale :**
   - Vue d’ensemble de la performance et des risques du portefeuille.

4. **Base de données :**
   - Création d'une base de données à partir des données brutes fournies.
   - Automatisation du traitement des données via VBA et SQL.
   - Requêtes démontrant le fonctionnement de la base.

5. **Documentation complète :**
   - Manuel utilisateur détaillé pour le client.
   - Note technique expliquant l'architecture et la logique du code.

---

## Architecture du Projet

Le projet est organisé en deux parties principales :

### 1. Base de Données (*fichier 1*)
- **Objectif :** Centralisation et structuration des données pour une utilisation efficace.
- **Étapes :**
  - Un script VBA traite les fichiers bruts et génère un fichier nettoyé.
  - Un script *Runs* séléctionne les actions selon un critère basé sur leurs runs
- **Résultats attendus :**
  - Une base fonctionnelle contenant les actions et obligations du portefeuille.

### 2. Outil de Gestion d’Investissement (*fichier 2*)
- **Objectif :** Permettre au client de consulter et analyser son portefeuille à tout moment.
- **Caractéristiques :**
  - Implémentation d’une classe VBA pour gérer la logique métier et les calculs.
  - Interaction avec les données nettoyées ou la base de données.
  - Génération de rapports et affichage des métriques demandées.

---

## Contraintes et Stratégies
1. **Allocation initiale :**
   - Au moins 50 % du portefeuille en actions, réparties sur au moins 5 produits différents.
2. **Obligations :**
   - 10 obligations fictives, dont 50 % versent un coupon semi-annuel.
3. **Données :**
   - Les données des actions sont extraites des fichiers bruts fournis.
   - Les obligations sont définies selon le modèle enseigné.

---
