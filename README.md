
# Insurance Brokerage Analytics

Projet Python de bout en bout consacré à l’analyse de la performance commerciale,
à la segmentation des clients et au suivi des délais de paiement des commissions
dans le secteur de l’assurance.

## Problématique

Comment analyser la performance d’un portefeuille commercial, identifier les
clients à forte valeur et repérer les opérations présentant un risque de retard
de paiement des commissions ?

## Contenu

- nettoyage et contrôle qualité ;
- indicateurs clés de performance ;
- analyse temporelle ;
- performance par compagnie et branche ;
- concentration du portefeuille ;
- segmentation RFM avec K-means ;
- classification des retards supérieurs à 60 jours ;
- dashboard Streamlit ;
- recommandations décisionnelles.

## Technologies

- Python
- pandas et NumPy
- Plotly et Matplotlib
- scikit-learn
- Streamlit
- Google Colab

## Données

Les données ont été anonymisées et transformées. Les noms de clients, les
numéros de police, les compagnies, les dates exactes et les montants exacts
ne sont pas publiés.

Une date de paiement non renseignée ne doit pas être interprétée automatiquement
comme un impayé.

## Structure recommandée

```text
insurance-brokerage-analytics/
├── data/
│   └── BASE_COMMERCIALE_PUBLIQUE_ANONYMISEE.xlsx
├── notebook/
│   └── Insurance_Brokerage_Analytics_Colab.ipynb
├── outputs/
├── streamlit_app.py
├── requirements.txt
└── README.md
```

## Résultats

Les résultats détaillés sont générés automatiquement par le notebook dans le
dossier `outputs`.

## Limites

- données anonymisées et montants transformés ;
- année 2026 partielle ;
- dates de paiement parfois absentes ;
- résultats prédictifs à utiliser comme aide à la décision.
