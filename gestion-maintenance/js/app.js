// Variables globales
let currentData = [];
let filteredData = [];
let currentWorkbook = null;
let dataTable = null;

// Initialisation
document.addEventListener('DOMContentLoaded', function() {
    initializeEventListeners();
    
    // Essayer de charger le fichier Excel par défaut si présent
    loadDefaultFile();
});

// Initialisation des écouteurs d'événements
function initializeEventListeners() {
    // Upload de fichier Excel
    document.getElementById('excelFile').addEventListener('change', handleFileUpload);
    
    // Changement de feuille
    document.getElementById('sheetSelect').addEventListener('change', function() {
        if (currentWorkbook) {
            const sheetName = this.value;
            loadSheetData(sheetName);
        }
    });
    
    // Filtres
    document.getElementById('wilayaFilter').addEventListener('change', applyFilters);
    document.getElementById('typeFilter').addEventListener('change', applyFilters);
    document.getElementById('technicienFilter').addEventListener('change', applyFilters);
    document.getElementById('statutFilter').addEventListener('change', applyFilters);
    
    // Boutons
    document.getElementById('resetFilters').addEventListener('click', resetFilters);
    document.getElementById('exportCSV').addEventListener('click', exportToCSV);
}

// Charger le fichier par défaut s'il existe
function loadDefaultFile() {
    fetch('data/gestion-maintenace.xlsx')
        .then(response => {
            if (response.ok) {
                return response.blob();
            }
            throw new Error('Fichier par défaut non trouvé');
        })
        .then(blob => {
            // Simuler un upload de fichier
            const file = new File([blob], 'gestion-maintenace.xlsx', { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
            readExcelFile(file);
        })
        .catch(error => {
            console.log('Charger un fichier manuellement:', error.message);
        });
}

// Gérer l'upload de fichier
function handleFileUpload(e) {
    const file = e.target.files[0];
    if (!file) return;
    
    // Mettre à jour le nom du fichier affiché
    document.getElementById('fileName').textContent = file.name;
    
    readExcelFile(file);
}

// Lire le fichier Excel
function readExcelFile(file) {
    const reader = new FileReader();
    
    reader.onload = function(e) {
        const data = new Uint8Array(e.target.result);
        try {
            currentWorkbook = XLSX.read(data, { type: 'array' });
            
            // Afficher les feuilles disponibles
            const sheetNames = currentWorkbook.SheetNames;
            console.log('Feuilles disponibles:', sheetNames);
            
            // Charger la feuille sélectionnée par défaut
            const defaultSheet = document.getElementById('sheetSelect').value;
            loadSheetData(defaultSheet);
            
        } catch (error) {
            alert('Erreur lors de la lecture du fichier Excel: ' + error.message);
            console.error(error);
        }
    };
    
    reader.onerror = function() {
        alert('Erreur lors de la lecture du fichier');
    };
    
    reader.readAsArrayBuffer(file);
}

// Charger les données d'une feuille spécifique
function loadSheetData(sheetName) {
    if (!currentWorkbook || !currentWorkbook.Sheets[sheetName]) {
        alert('Feuille non trouvée: ' + sheetName);
        return;
    }
    
    const worksheet = currentWorkbook.Sheets[sheetName];
    
    // Convertir en JSON avec gestion des en-têtes
    const jsonData = XLSX.utils.sheet_to_json(worksheet, { defval: '' });
    
    console.log('Données brutes de la feuille', sheetName + ':', jsonData);
    
    // Normaliser les données
    currentData = normalizeData(jsonData);
    
    // Appliquer les filtres courants
    applyFilters();
    
    // Mettre à jour le titre de la page
    document.querySelector('header h1').textContent = `📊 Système de Gestion de Maintenance - ${sheetName}`;
}

// Normaliser les données (gérer les différentes structures de colonnes)
function normalizeData(data) {
    if (!data || data.length === 0) return [];
    
    // Tenter de détecter la structure des colonnes
    const firstRow = data[0];
    const normalizedData = [];
    
    data.forEach(row => {
        // Déterminer la colonne d'équipement
        const equipement = row['équipement'] || row['Equipement'] || row['EQUIPEMENT'] || '';
        
        // Seulement inclure les lignes avec un équipement spécifié
        if (equipement && equipement.trim() !== '') {
            const normalizedRow = {
                equipement: equipement,
                marque: row['marque'] || row['Marque'] || row['MARQUE'] || '',
                inventaire: row['inventaire'] || row['Inventaire'] || row['INVENTAIRE'] || '',
                serie: row['n° série'] || row['N° série'] || row['N° Série'] || row['serie'] || '',
                etablissement: row['établissement'] || row['Etablissement'] || row['ETABLISSEMENT'] || '',
                type: detectType(row['établissement'] || ''),
                panne: row['panne'] || row['Panne'] || row['PANNE'] || '',
                technicien: row['technicien'] || row['Technicien'] || row['TECHNICIEN'] || '',
                statut: detectStatut(row),
                date: row['date'] || row['Date'] || row['DATE'] || ''
            };
            
            normalizedData.push(normalizedRow);
        }
    });
    
    return normalizedData;
}

// Détecter le type d'établissement
function detectType(etablissement) {
    const etablissementStr = etablissement.toString().toUpperCase();
    
    if (etablissementStr.includes('EP') || etablissementStr.includes('PRIMAIRE')) {
        return 'EP';
    } else if (etablissementStr.includes('CEM')) {
        return 'CEM';
    } else if (etablissementStr.includes('LYCEE') || etablissementStr.includes('LYCÉE')) {
        return 'Lycée';
    } else if (etablissementStr.includes('DIRECTION')) {
        return 'Direction';
    }
    
    return 'Autre';
}

// Détecter le statut de maintenance
function detectStatut(row) {
    // Chercher les colonnes de statut (rec, re, nr)
    if (row['rec'] === 1 || row['REC'] === 1 || row['Reçu'] === 1) {
        return 'rec';
    } else if (row['re'] === 1 || row['RE'] === 1 || row['Réparé'] === 1) {
        return 're';
    } else if (row['nr'] === 1 || row['NR'] === 1 || row['Non réparé'] === 1) {
        return 'nr';
    }
    
    // Essayer de détecter depuis le texte
    const text = JSON.stringify(row).toLowerCase();
    if (text.includes('reçu') || text.includes('recu')) {
        return 'rec';
    } else if (text.includes('réparé') || text.includes('reparé') || text.includes('repar')) {
        return 're';
    } else if (text.includes('non réparé') || text.includes('non reparé') || text.includes('nr')) {
        return 'nr';
    }
    
    return 'rec'; // Par défaut
}

// Appliquer les filtres
function applyFilters() {
    if (currentData.length === 0) return;
    
    // Récupérer les valeurs des filtres
    const wilayaFilter = document.getElementById('wilayaFilter').value;
    const typeFilter = document.getElementById('typeFilter').value;
    const technicienFilter = document.getElementById('technicienFilter').value;
    const statutFilter = document.getElementById('statutFilter').value;
    
    // Filtrer les données
    filteredData = currentData.filter(row => {
        // Filtre par wilaya (toujours El Bayadh pour l'instant)
        if (wilayaFilter !== 'all' && wilayaFilter !== 'El Bayadh') {
            return false;
        }
        
        // Filtre par type d'établissement
        if (typeFilter !== 'all' && row.type !== typeFilter) {
            return false;
        }
        
        // Filtre par technicien
        if (technicienFilter !== 'all' && row.technicien !== technicienFilter) {
            return false;
        }
        
        // Filtre par statut
        if (statutFilter !== 'all' && row.statut !== statutFilter) {
            return false;
        }
        
        return true;
    });
    
    // Mettre à jour l'affichage
    updateDisplay();
}

// Réinitialiser tous les filtres
function resetFilters() {
    document.getElementById('wilayaFilter').value = 'El Bayadh';
    document.getElementById('typeFilter').value = 'all';
    document.getElementById('technicienFilter').value = 'all';
    document.getElementById('statutFilter').value = 'all';
    
    applyFilters();
}

// Mettre à jour l'affichage (tableau et statistiques)
function updateDisplay() {
    updateStats();
    updateCharts();
    updateTable();
}

// Mettre à jour les statistiques
function updateStats() {
    const total = filteredData.length;
    const recus = filteredData.filter(d => d.statut === 'rec').length;
    const reparés = filteredData.filter(d => d.statut === 're').length;
    const nonRepares = filteredData.filter(d => d.statut === 'nr').length;
    
    // Mettre à jour les compteurs
    document.getElementById('totalEquipments').textContent = total;
    document.getElementById('recusCount').textContent = recus;
    document.getElementById('reparésCount').textContent = reparés;
    document.getElementById('nonReparesCount').textContent = nonRepares;
}

// Mettre à jour les graphiques
function updateCharts() {
    updateStatusChart();
    updateTypeChart();
}

// Graphique des statuts
function updateStatusChart() {
    const ctx = document.getElementById('statusChart').getContext('2d');
    
    // Détruire le graphique existant s'il existe
    if (window.statusChart instanceof Chart) {
        window.statusChart.destroy();
    }
    
    const statusCounts = {
        'Reçus': filteredData.filter(d => d.statut === 'rec').length,
        'Réparés': filteredData.filter(d => d.statut === 're').length,
        'Non réparés': filteredData.filter(d => d.statut === 'nr').length
    };
    
    window.statusChart = new Chart(ctx, {
        type: 'doughnut',
        data: {
            labels: Object.keys(statusCounts),
            datasets: [{
                data: Object.values(statusCounts),
                backgroundColor: [
                    '#4299e1', // Bleu pour reçus
                    '#48bb78', // Vert pour réparés
                    '#f56565'  // Rouge pour non réparés
                ],
                borderWidth: 2,
                borderColor: '#fff'
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: {
                    position: 'bottom',
                    labels: {
                        padding: 20,
                        font: {
                            size: 14
                        }
                    }
                },
                title: {
                    display: true,
                    text: 'Répartition par Statut',
                    font: {
                        size: 16,
                        weight: 'bold'
                    }
                }
            }
        }
    });
}

// Graphique par type d'établissement
function updateTypeChart() {
    const ctx = document.getElementById('typeChart').getContext('2d');
    
    // Détruire le graphique existant s'il existe
    if (window.typeChart instanceof Chart) {
        window.typeChart.destroy();
    }
    
    const typeCounts = {
        'EP': filteredData.filter(d => d.type === 'EP').length,
        'CEM': filteredData.filter(d => d.type === 'CEM').length,
        'Lycée': filteredData.filter(d => d.type === 'Lycée').length,
        'Direction': filteredData.filter(d => d.type === 'Direction').length,
        'Autre': filteredData.filter(d => d.type === 'Autre').length
    };
    
    window.typeChart = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: Object.keys(typeCounts),
            datasets: [{
                label: 'Nombre d\'équipements',
                data: Object.values(typeCounts),
                backgroundColor: [
                    '#667eea', // EP
                    '#764ba2', // CEM
                    '#f687b3', // Lycée
                    '#f6ad55', // Direction
                    '#cbd5e0'  // Autre
                ],
                borderWidth: 1
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            scales: {
                y: {
                    beginAtZero: true,
                    title: {
                        display: true,
                        text: 'Nombre d\'équipements'
                    }
                },
                x: {
                    title: {
                        display: true,
                        text: 'Type d\'établissement'
                    }
                }
            },
            plugins: {
                legend: {
                    display: false
                },
                title: {
                    display: true,
                    text: 'Répartition par Type d\'Établissement',
                    font: {
                        size: 16,
                        weight: 'bold'
                    }
                }
            }
        }
    });
}

// Mettre à jour le tableau
function updateTable() {
    const tableBody = document.getElementById('tableBody');
    tableBody.innerHTML = '';
    
    // Trier les données par date (si disponible)
    const sortedData = [...filteredData].sort((a, b) => {
        if (!a.date) return 1;
        if (!b.date) return -1;
        return new Date(b.date) - new Date(a.date);
    });
    
    // Remplir le tableau
    sortedData.forEach(row => {
        const tr = document.createElement('tr');
        
        // Badge de statut
        let statutText = '';
        let statutClass = '';
        switch(row.statut) {
            case 'rec':
                statutText = 'Reçu';
                statutClass = 'statut-rec';
                break;
            case 're':
                statutText = 'Réparé';
                statutClass = 'statut-re';
                break;
            case 'nr':
                statutText = 'Non réparé';
                statutClass = 'statut-nr';
                break;
            default:
                statutText = row.statut;
                statutClass = 'statut-rec';
        }
        
        tr.innerHTML = `
            <td>${escapeHtml(row.equipement)}</td>
            <td>${escapeHtml(row.marque)}</td>
            <td>${escapeHtml(row.inventaire)}</td>
            <td>${escapeHtml(row.serie)}</td>
            <td>${escapeHtml(row.etablissement)}</td>
            <td>${escapeHtml(row.type)}</td>
            <td>${escapeHtml(row.panne)}</td>
            <td>${escapeHtml(row.technicien)}</td>
            <td><span class="statut-badge ${statutClass}">${statutText}</span></td>
            <td>${escapeHtml(row.date)}</td>
        `;
        
        tableBody.appendChild(tr);
    });
    
    // Initialiser ou re-initialiser DataTables
    if (dataTable) {
        dataTable.destroy();
    }
    
    dataTable = $('#dataTable').DataTable({
        language: {
            url: '//cdn.datatables.net/plug-ins/1.13.4/i18n/fr-FR.json'
        },
        pageLength: 10,
        lengthMenu: [5, 10, 25, 50, 100],
        order: [[9, 'desc']], // Trier par date décroissante
        dom: 'Bfrtip',
        buttons: [
            'copy', 'csv', 'excel', 'pdf', 'print'
        ],
        responsive: true
    });
}

// Échapper les caractères HTML pour la sécurité
function escapeHtml(text) {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
}

// Exporter en CSV
function exportToCSV() {
    if (filteredData.length === 0) {
        alert('Aucune donnée à exporter');
        return;
    }
    
    // Convertir en CSV
    const headers = ['Équipement', 'Marque', 'Inventaire', 'N° Série', 'Établissement', 'Type', 'Panne', 'Technicien', 'Statut', 'Date'];
    const csvRows = [
        headers.join(','),
        ...filteredData.map(row => [
            `"${row.equipement}"`,
            `"${row.marque}"`,
            `"${row.inventaire}"`,
            `"${row.serie}"`,
            `"${row.etablissement}"`,
            `"${row.type}"`,
            `"${row.panne}"`,
            `"${row.technicien}"`,
            `"${row.statut}"`,
            `"${row.date}"`
        ].join(','))
    ];
    
    const csvString = csvRows.join('\n');
    const blob = new Blob([csvString], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    
    // Créer un lien de téléchargement
    const link = document.createElement('a');
    link.href = url;
    link.download = `maintenance_${document.getElementById('sheetSelect').value}_${new Date().toISOString().split('T')[0]}.csv`;
    link.click();
    
    // Nettoyer
    URL.revokeObjectURL(url);
}