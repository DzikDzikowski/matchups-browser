let database = {};
let myChampionPool = JSON.parse(localStorage.getItem('lolMyPool')) || []; // <--- TUTAJ: zmienione z ['yone', 'sion'...] na puste []
let mainRole = "";

let currentEditingEnemy = "";
let currentEditingRole = "";

function init() {
    allLeagueChampions.forEach(champ => {
        database[champ] = { "Toplane": [], "Jungle": [], "Midlane": [], "Botlane": [], "Support": [] };
    });
    mainRole = "";

    renderMainRole();
    populatePoolDropdown('');
    renderPoolTags();
    setupEventListeners();
    displayResults('');
}

function calculateMainRole() {
    let roleStats = { "Toplane": 0, "Jungle": 0, "Midlane": 0, "Botlane": 0, "Support": 0 };
    
    for (const enemy in database) {
        for (const role in database[enemy]) {
            roleStats[role] += database[enemy][role].length;
        }
    }

    let maxScore = 0;
    let currentMain = "Toplane";
    for (const [role, score] of Object.entries(roleStats)) {
        if (score > maxScore) {
            maxScore = score;
            currentMain = role;
        }
    }
    mainRole = currentMain;
}

function saveData() {
    calculateMainRole();
    renderMainRole();
}

function renderMainRole() {
    const container = document.getElementById('mainRoleDisplay');
    if (mainRole) {
        container.innerHTML = `<span class="role-badge">Twoja główna rola: ${ROLE_ICONS[mainRole]}</span>`;
    } else {
        container.innerHTML = '';
    }
}

function normalizeChampionName(name) {
    let trimmed = name.trim();
    let lower = trimmed.toLowerCase();
    const corrections = {
        "malaphite": "Malphite", "malaphte": "Malphite", "renkton": "Renekton",
        "orn": "Ornn", "zahhen": "Zaahen", "k'sante": "K'Sante"
    };
    
    let corrected = corrections[lower] || trimmed;
    const properName = allLeagueChampions.find(c => c.toLowerCase() === corrected.toLowerCase());
    return properName || corrected; 
}

function getImageUrl(champName) {
    let name = champName;
    const exceptions = {
        "wukong": "MonkeyKing", "dr. mundo": "DrMundo", "k'sante": "KSante",
        "cho'gath": "Chogath", "rek'sai": "RekSai", "bel'veth": "Belveth",
        "kai'sa": "Kaisa", "kha'zix": "Khazix", "leblanc": "Leblanc",
        "vel'koz": "Velkoz", "kog'maw": "KogMaw", "nunu & willump": "Nunu",
        "renata glasc": "Renata", "jarvan iv": "JarvanIV"
    };

    const lowerName = name.toLowerCase().trim();
    if (exceptions[lowerName]) {
        name = exceptions[lowerName];
    } else {
        name = name.split(/[\s'.]+/).map(part => {
            if(!part) return "";
            return part.charAt(0).toUpperCase() + part.slice(1).toLowerCase();
        }).join('');
    }
    return `http://ddragon.leagueoflegends.com/cdn/16.8.1/img/champion/${name}.png`;
}

function processExcelFile(arrayBuffer) {
    const workbook = XLSX.read(arrayBuffer, {type: 'array'});
    
    database = {};
    allLeagueChampions.forEach(champ => {
        database[champ] = { "Toplane": [], "Jungle": [], "Midlane": [], "Botlane": [], "Support": [] };
    });

    workbook.SheetNames.forEach(sheetName => {
        const role = ROLES.find(r => r.toLowerCase() === sheetName.toLowerCase());
        if (!role) return;

        const sheet = workbook.Sheets[sheetName];
        const jsonData = XLSX.utils.sheet_to_json(sheet, {header: 1});

        for (let i = 1; i < jsonData.length; i++) {
            const row = jsonData[i];
            if (!row || !row[0]) continue; 

            const enemyRaw = normalizeChampionName(row[0]);

            if (!database[enemyRaw]) database[enemyRaw] = { "Toplane": [], "Jungle": [], "Midlane": [], "Botlane": [], "Support": [] };

            const countersSet = new Set();
            for (let j = 1; j < row.length; j++) {
                const cell = row[j];
                if (cell && typeof cell === 'string' && cell.trim() !== '') {
                    countersSet.add(normalizeChampionName(cell));
                }
            }
            database[enemyRaw][role] = Array.from(countersSet);
        }
    });

    saveData();
}

function exportDatabaseToExcel() {
    const workbook = XLSX.utils.book_new();

    ROLES.forEach(role => {
        const sheetRows = [];
        
        sheetRows.push(["Przeciwnik", "Kontra 1", "Kontra 2", "Kontra 3", "Kontra 4", "Kontra 5"]);

        for (const enemy in database) {
            const counters = database[enemy][role] || [];
            
            if (counters.length > 0) {
                const row = [enemy, ...counters];
                sheetRows.push(row);
            }
        }

        const worksheet = XLSX.utils.aoa_to_sheet(sheetRows);
        XLSX.utils.book_append_sheet(workbook, worksheet, role);
    });

    XLSX.writeFile(workbook, "lol_matchups_updated.xlsx");
}

function populatePoolDropdown(filterText) {
    const dropdown = document.getElementById('champDropdown');
    dropdown.innerHTML = '';
    const searchLower = filterText.toLowerCase();
    const filtered = allLeagueChampions.filter(champ => champ.toLowerCase().includes(searchLower));

    if (filtered.length === 0) {
        dropdown.innerHTML = '<div class="no-suggestions" style="padding:10px; color:#777;">Nie znaleziono postaci</div>';
        return;
    }

    filtered.forEach(champName => {
        const item = document.createElement('div');
        item.className = 'dropdown-item';
        item.innerHTML = `<img src="${getImageUrl(champName)}" onerror="this.style.opacity='0'"><span>${champName}</span>`;
        item.addEventListener('click', () => {
            addChampionToPool(champName.toLowerCase());
            dropdown.classList.remove('show');
            document.getElementById('poolSearchInput').value = ''; 
        });
        dropdown.appendChild(item);
    });
}

function renderPoolTags() {
    const container = document.getElementById('tagsContainer');
    container.innerHTML = '';
    myChampionPool.forEach(champ => {
        const tag = document.createElement('div');
        tag.className = 'champ-tag';
        const properName = allLeagueChampions.find(c => c.toLowerCase() === champ) || champ;
        tag.innerHTML = `<span>${properName}</span><span class="remove-btn" onclick="removeChampionFromPool('${champ}')">&times;</span>`;
        container.appendChild(tag);
    });
    localStorage.setItem('lolMyPool', JSON.stringify(myChampionPool));
}

function addChampionToPool(name) {
    if (name && !myChampionPool.includes(name)) {
        myChampionPool.push(name);
        renderPoolTags();
        displayResults(document.getElementById('searchInput').value);
    }
}

window.removeChampionFromPool = function(champName) {
    myChampionPool = myChampionPool.filter(c => c !== champName);
    renderPoolTags();
    displayResults(document.getElementById('searchInput').value);
}

window.openAddCounterModal = function(enemy, role) {
    currentEditingEnemy = enemy;
    currentEditingRole = role;
    
    const modal = document.getElementById('addCounterModal');
    const title = document.getElementById('modalTitle');
    const searchInput = document.getElementById('modalSearchInput');
    
    title.innerHTML = `Dodaj kontrę na: <strong style="color:var(--accent-color);">${enemy}</strong> (${ROLE_ICONS[role] || role})`;
    searchInput.value = '';
    
    populateModalChampList('');
    modal.classList.add('show');
}

window.closeModal = function() {
    document.getElementById('addCounterModal').classList.remove('show');
}

function populateModalChampList(filterText) {
    const container = document.getElementById('modalChampList');
    container.innerHTML = '';
    
    const searchLower = filterText.toLowerCase();
    
    const filtered = allLeagueChampions.filter(champ => {
        const matchesSearch = champ.toLowerCase().includes(searchLower);
        const rolesForChamp = championRoles[champ] || [];
        const matchesRole = rolesForChamp.includes(currentEditingRole);
        return matchesSearch && matchesRole;
    });

    if (filtered.length === 0) {
        container.innerHTML = '<div style="padding:15px; color:#777; text-align:center; font-style:italic;">Brak pasujących postaci dla tej roli</div>';
        return;
    }

    filtered.forEach(champName => {
        const currentCounters = database[currentEditingEnemy]?.[currentEditingRole] || [];
        const isAlreadyCounter = currentCounters.includes(champName);

        const item = document.createElement('div');
        item.className = 'dropdown-item';
        if(isAlreadyCounter) {
            item.style.opacity = '0.4';
            item.style.cursor = 'not-allowed';
        }
        
        item.innerHTML = `
            <img src="${getImageUrl(champName)}" onerror="this.style.opacity='0'">
            <span style="flex-grow:1;">${champName}</span>
            ${isAlreadyCounter ? '<span style="font-size:12px; color:#777;">Już dodany</span>' : ''}
        `;
        
        if(!isAlreadyCounter) {
            item.addEventListener('click', () => {
                addCounterToDatabase(currentEditingEnemy, currentEditingRole, champName);
                closeModal();
            });
        }
        container.appendChild(item);
    });
}

function addCounterToDatabase(enemy, role, counterName) {
    if (!database[enemy]) {
        database[enemy] = { "Toplane": [], "Jungle": [], "Midlane": [], "Botlane": [], "Support": [] };
    }
    if (!database[enemy][role]) {
        database[enemy][role] = [];
    }
    
    if (!database[enemy][role].includes(counterName)) {
        database[enemy][role].push(counterName);
        saveData();
        displayResults(document.getElementById('searchInput').value);
    }
}

window.removeCounterFromDatabase = function(enemy, role, counterName) {
    if (database[enemy] && database[enemy][role]) {
        database[enemy][role] = database[enemy][role].filter(c => c !== counterName);
        saveData();
        displayResults(document.getElementById('searchInput').value);
    }
}

function displayResults(query) {
    const container = document.getElementById('resultsContainer');
    container.innerHTML = '';

    if (!query) {
        return;
    }

    const searchLower = query.toLowerCase();
    const matchKey = Object.keys(database).find(enemy => enemy.toLowerCase() === searchLower) || 
                     Object.keys(database).find(enemy => enemy.toLowerCase().includes(searchLower));

    if (!matchKey) {
        container.innerHTML = '<div class="no-results" style="text-align:center;">Nie znaleziono takiego championa.</div>';
        return;
    }

    const header = document.createElement('div');
    header.className = 'results-header';
    header.innerHTML = `Przeciwnik: <strong style="color:var(--accent-color);">${matchKey}</strong>`;
    container.appendChild(header);

    let displayOrder = [];
    if(mainRole && ROLES.includes(mainRole)) {
        displayOrder.push(mainRole);
    }
    ROLES.forEach(r => {
        if(!displayOrder.includes(r)) displayOrder.push(r);
    });

    const enemyData = database[matchKey];
    const availableRoles = championRoles[matchKey] || [];

    displayOrder.filter(role => availableRoles.includes(role)).forEach(role => {
        const counters = enemyData[role] || [];
        
        const section = document.createElement('div');
        section.className = 'role-section';
        
        const roleHeader = document.createElement('div');
        roleHeader.className = 'role-header';
        roleHeader.innerHTML = `
            <span class="role-title">${ROLE_ICONS[role] || role}</span>
            <button class="add-counter-btn" onclick="openAddCounterModal('${matchKey}', '${role}')" title="Dodaj kontrę dla tej roli">+</button>
        `;
        section.appendChild(roleHeader);

        if (counters.length === 0) {
            const emptyMsg = document.createElement('div');
            emptyMsg.className = 'no-results-role';
            emptyMsg.innerText = "Brak przypisanych kontr dla tej roli.";
            section.appendChild(emptyMsg);
        } else {
            const grid = document.createElement('div');
            grid.className = 'counters-grid';

            const poolCounters = [];
            const _otherCounters = [];

            counters.forEach(counter => {
                const isPoolChampion = myChampionPool.includes(counter.toLowerCase().trim());
                if (isPoolChampion) poolCounters.push(counter);
                else _otherCounters.push(counter);
            });

            const sortedCounters = [...poolCounters, ..._otherCounters];

            sortedCounters.forEach(counter => {
                const card = document.createElement('div');
                card.className = 'champion-card';
                
                const isPoolChampion = myChampionPool.includes(counter.toLowerCase().trim());
                if (isPoolChampion) card.classList.add('pool-champion');
                
                card.innerHTML = `
                    <span class="remove-counter-btn" onclick="removeCounterFromDatabase('${matchKey}', '${role}', '${counter}')" title="Usuń kontrę">&times;</span>
                    <img class="champion-icon" src="${getImageUrl(counter)}" onerror="this.style.display='none'">
                    <div class="champion-name">${counter}</div>
                `;
                grid.appendChild(card);
            });
            section.appendChild(grid);
        }

        container.appendChild(section);
    });
}

function setupEventListeners() {
    document.getElementById('searchInput').addEventListener('input', (e) => {
        displayResults(e.target.value);
    });

    const poolSearchInput = document.getElementById('poolSearchInput');
    const dropdown = document.getElementById('champDropdown');
    poolSearchInput.addEventListener('input', (e) => {
        populatePoolDropdown(e.target.value);
        dropdown.classList.add('show');
    });
    poolSearchInput.addEventListener('focus', () => {
        populatePoolDropdown(poolSearchInput.value);
        dropdown.classList.add('show');
    });

    document.getElementById('modalSearchInput').addEventListener('input', (e) => {
        populateModalChampList(e.target.value);
    });

    document.addEventListener('click', (e) => {
        if (!poolSearchInput.contains(e.target) && !dropdown.contains(e.target)) {
            dropdown.classList.remove('show');
        }
        const modal = document.getElementById('addCounterModal');
        if (e.target === modal) {
            closeModal();
        }
    });

    document.getElementById('excelFileInput').addEventListener('change', (e) => {
        const file = e.target.files[0];
        if (file) {
            const reader = new FileReader();
            reader.onload = function(event) {
                try {
                    const data = new Uint8Array(event.target.result);
                    processExcelFile(data);
                    alert('Baza danych z pliku Excel została wgrana pomyślnie!');
                    displayResults(document.getElementById('searchInput').value);
                } catch (error) {
                    alert('Wystąpił błąd podczas odczytu pliku.');
                    console.error(error);
                }
            };
            reader.readAsArrayBuffer(file);
        }
    });

    document.getElementById('downloadExcelBtn').addEventListener('click', () => {
        exportDatabaseToExcel();
    });
}

window.onload = init;