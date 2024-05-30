import pandas as pd
import re

# Fonction pour extraire les détails du filtre
def extract_filter_details(filter_name):
    details = {
        'Ethertype': 'IPv4'  # valeur par défaut, sera ajustée selon le besoin
    }
    if 'TCP' in filter_name:
        details['Protocol'] = 'TCP'
    elif 'UDP' in filter_name:
        details['Protocol'] = 'UDP'
    
    # Extraction des plages de ports
    ports = filter_name.split('_')[0]
    port_parts = re.split(r'\s*-\s*|\s*\u2026\s*', ports)

    if len(port_parts) == 2:
        details['destination_from_port'] = port_parts[0]
        details['destination_to_port'] = port_parts[1]
    elif len(port_parts) == 1:
        details['destination_from_port'] = port_parts[0]
        details['destination_to_port'] = port_parts[0]
    else:
        details['destination_from_port'] = None
        details['destination_to_port'] = None

    return details

# Chargement du fichier Excel A MODIFIER
df = pd.read_excel('/Users/astro14/SynologyDrive/1.ADG Consulting/2.Clients/TV5 monde/Intégration/ContratFiltres/script/vrf_excel_file.xlsx')

# Suppression des lignes en doublon
df.drop_duplicates(inplace=True)

# Renommage de la colonne 'VRF' en 'Tenant'
df.rename(columns={'VRF': 'Tenant'}, inplace=True)

# Changer le suffixe '_VRF' par '_TN' dans les valeurs de la colonne 'Tenant'
df['Tenant'] = df['Tenant'].astype(str).str.replace('_VRF', '_TN')

# Création du DataFrame des filtres uniques avec le Tenant associé
unique_filters = df[['Ports', 'Tenant']].drop_duplicates().rename(columns={'Ports': 'Filter'})

# Appliquer la fonction pour extraire les détails des filtres
filter_details = unique_filters['Filter'].apply(lambda x: pd.Series(extract_filter_details(x)))
unique_filters_detailed = pd.concat([unique_filters, filter_details], axis=1)

# Convertir les valeurs des colonnes en minuscules
unique_filters_detailed['Ethertype'] = unique_filters_detailed['Ethertype'].astype(str).str.lower()
unique_filters_detailed['Protocol'] = unique_filters_detailed['Protocol'].astype(str).str.lower()

# Ajouter la colonne "Filter entries"
unique_filters_detailed['Filter entries'] = unique_filters_detailed['Filter']

# Ajouter la colonne "Stateful"
unique_filters_detailed['Stateful'] = unique_filters_detailed['Protocol'].map({'tcp': True, 'udp': False})

# Réorganiser les colonnes pour mettre "Tenant" en premier
cols = ['Tenant', 'Filter', 'Filter entries', 'Ethertype', 'Protocol', 'destination_from_port', 'destination_to_port', 'Stateful']
unique_filters_detailed = unique_filters_detailed[cols]

# Pivoter le DataFrame initial pour obtenir une vue des contrats
df['Index'] = df.groupby(['EPG source', 'EPG destination']).cumcount()
pivot_columns = ['EPG source', 'EPG destination', 'Tenant'] + [f'Filter {i}' for i in range(int(df['Index'].max()) + 1)]
df_pivoted = df.pivot_table(index=['EPG source', 'EPG destination', 'Tenant'], columns='Index', values='Ports', aggfunc='first').reset_index()
df_pivoted.columns = pivot_columns
df_pivoted['Contract Name'] = df_pivoted['EPG source'] + '...' + df_pivoted['EPG destination'] + '_Ct'

# Ajouter la colonne "Subjects"
df_pivoted['Subject'] = df_pivoted['Contract Name'].astype(str).str.replace('_Ct', '_Sbj')

# Réorganiser pour inclure 'Subjects' entre 'Contract Name' et 'EPG source'
cols_pivot = ['Tenant', 'Contract Name', 'Subject', 'EPG source', 'EPG destination'] + [col for col in df_pivoted.columns if col not in ['Tenant', 'Contract Name', 'Subject', 'EPG source', 'EPG destination']]
df_pivoted = df_pivoted[cols_pivot]

# Extraire les données pour l'onglet 'Contrat_EPGs'
df_contract_epgs = df_pivoted[['Tenant', 'Contract Name', 'EPG source', 'EPG destination']].copy()

# Ajouter la colonne 'AppProfile' en utilisant .loc pour éviter SettingWithCopyWarning
df_contract_epgs.loc[:, 'AppProfile'] = df_contract_epgs['Tenant'].map({
    'XDA_ADM_TN': 'APP_ADM',
    'XDA_DATA_TN': 'APP_DATA'
})

# Réorganiser les colonnes pour mettre 'AppProfile' à la place de 'Subject'
cols_contract_epgs = ['Tenant', 'Contract Name', 'AppProfile', 'EPG source', 'EPG destination']
df_contract_epgs = df_contract_epgs[cols_contract_epgs]

# Extraire les données pour l'onglet 'ContratToFilters'
df_contrat_to_filters = df_pivoted.drop(columns=['EPG source', 'EPG destination'])

# Enregistrer le DataFrame dans un nouveau fichier Excel A MODIFIER
output_path = 'ListeContratEtFiltresMultiOnglets.xlsx'
with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
    df_pivoted.to_excel(writer, sheet_name='Contracts_Filters_EPGs', index=False)
    unique_filters_detailed.to_excel(writer, sheet_name='Unique_Filters', index=False)
    df_contract_epgs.to_excel(writer, sheet_name='Contrat_EPGs', index=False)
    df_contrat_to_filters.to_excel(writer, sheet_name='ContratToFilters', index=False)
