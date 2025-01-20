# Gestion de Portefeuille d'Investissement

## Description du Projet

Ce projet vise à développer un **outil de gestion de portefeuille d’investissement** destiné à un client ayant confié une somme initiale de **1 000 000 € au 1er janvier 2005**. L’outil permettra de suivre et d’analyser la composition, la valeur, la performance et les mesures de risque d’un portefeuille composé d’actions et d’obligations.  

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

### 1. Base de Données
- **Objectif :** Centraliser et structurer les données pour une utilisation efficace.
- **Étapes :**
  - Un script VBA traite les fichiers bruts et génère un fichier nettoyé.
  - Un script SQL crée une base de données et y importe les données nettoyées.
  - Deux requêtes SQL démontrent la fonctionnalité : une pour les obligations, une pour deux indices sur une période choisie.
- **Résultats attendus :**
  - Une base fonctionnelle contenant les actions et obligations du portefeuille.

### 2. Outil de Gestion d’Investissement
- **Objectif :** Permettre au client de consulter et analyser son portefeuille à tout moment.
- **Caractéristiques :**
  - Implémentation d’une classe VBA pour gérer la logique métier et les calculs.
  - Interaction avec les données nettoyées ou la base de données.
  - Génération de rapports et affichage des métriques demandées.

---

## Livrables

Le rendu final contient les fichiers suivants :
1. **Script VBA de traitement des données** (`data_processing.xlsm`) :
   - Nettoie et structure les données initiales.
   - Produit un fichier exploitable pour la base de données.
   
2. **Script SQL de création de la base** (`database_setup.sql`) :
   - Crée la base de données.
   - Importe les données traitées.
   - Contient deux requêtes démonstratives.

3. **Outil de gestion** (`portfolio_manager.xlsm`) :
   - Interface pour le suivi et l’analyse du portefeuille.

4. **Manuel utilisateur** (`manuel_utilisateur.pdf`) :
   - Guide étape par étape pour l’utilisation de l’outil.

5. **Note technique** (`note_technique.pdf`) :
   - Documentation détaillée pour un informaticien.
   - Explications sur la structure et le fonctionnement du code.

---

## Contraintes et Stratégies
1. **Allocation initiale :**
   - Au moins 50 % du portefeuille en actions, réparties sur au moins 5 produits différents.
2. **Obligations :**
   - 10 obligations fictives, dont 50 % versent un coupon semi-annuel.
3. **Données :**
   - Les données des actions sont extraites des fichiers bruts fournis.
   - Les obligations sont définies selon le modèle enseigné.
4. **Date limite de rendu :** Dimanche 28 avril à 18h00.

---

## Barème d'Évaluation

Le projet sera noté sur 20 points répartis comme suit :
- **5 points** : Respect du sujet et documentation.
- **5 points** : Implémentation SQL.
- **5 points** : Conception de la base de données.
- **5 points** : Fonctionnalité de l’outil.

---

## Utilisation de l'Outil

### Pré-requis
- Microsoft Excel (compatible avec les macros VBA).
- Serveur SQL (pour exécuter le script SQL).
- Données brutes disponibles localement.

### Instructions
1. **Préparation des données :**
   - Exécutez le script VBA pour traiter les données brutes.
   - Un fichier nettoyé sera généré.
2. **Création de la base de données :**
   - Importez le fichier nettoyé à l’aide du script SQL.
   - Exécutez les requêtes incluses pour valider la base.
3. **Suivi du portefeuille :**
   - Ouvrez l’outil VBA.
   - Suivez les instructions du manuel utilisateur pour consulter et analyser le portefeuille.

---

## Auteurs
Ce projet a été réalisé par :
- [Nom de Famille 1]
- [Nom de Famille 2]
- [Nom de Famille 3]

---

## Contact
Pour toute question, veuillez contacter l’équipe par email : [votre_email@example.com].
