document.addEventListener('DOMContentLoaded', () => {
    const addRowButton = document.getElementById('addRow');
    const addColumnButton = document.getElementById('addColumn');
    const dataTable = document.getElementById('dataTable');

    // --- CONTROLE DE PÁGINAS POR MÊS/ANO ---
    // Adiciona controles de navegação de mês
    const monthNames = [
        "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
        "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"
    ];
    let currentDate = new Date();
    let currentMonth = currentDate.getMonth();
    let currentYear = currentDate.getFullYear();

    // Cria barra de navegação de mês
    const navBar = document.createElement('div');
    navBar.style.display = 'flex';
    navBar.style.justifyContent = 'center';
    navBar.style.alignItems = 'center';
    navBar.style.gap = '10px';
    navBar.style.marginBottom = '20px';

    const prevBtn = document.createElement('button');
    prevBtn.textContent = '<';
    const nextBtn = document.createElement('button');
    nextBtn.textContent = '>';
    const monthLabel = document.createElement('span');
    monthLabel.style.fontWeight = 'bold';

    navBar.appendChild(prevBtn);
    navBar.appendChild(monthLabel);
    navBar.appendChild(nextBtn);

    // Insere a barra antes do primeiro .container
    document.body.insertBefore(navBar, document.querySelector('.container'));

    function updateMonthLabel() {
        monthLabel.textContent = `${monthNames[currentMonth]} / ${currentYear}`;
    }

    function getStorageKeyReceitas() {
        return `tableData_${currentYear}_${currentMonth+1}`;
    }
    function getStorageKeyDespesas() {
        return `tableDataExpenses_${currentYear}_${currentMonth+1}`;
    }

    // --- ATUALIZAÇÕES NAS FUNÇÕES PARA USAR CHAVES POR MÊS/ANO ---

    // --- Função utilitária para formatar valores em euro ---
    function formatCurrency(value) {
        return Number(value).toLocaleString('pt-PT', { style: 'currency', currency: 'EUR' });
    }

    // Função utilitária para converter texto de moeda para número
    function parseCurrencyText(text) {
        if (!text) return 0;
        // Remove tudo que não for número, vírgula, ponto ou sinal
        let cleaned = text.replace(/[^\d\.,-]/g, '').replace(',', '.');
        let num = parseFloat(cleaned);
        return isNaN(num) ? 0 : num;
    }

    // Receitas
    const loadTableData = () => {
        const savedData = JSON.parse(localStorage.getItem(getStorageKeyReceitas())) || { headers: [], rows: [] };
        const thead = dataTable.querySelector('thead tr');
        const tbody = dataTable.querySelector('tbody');

        // Populate headers
        thead.innerHTML = '';
        // Adiciona header vazio para coluna dos botões
        const thMove = document.createElement('th');
        thMove.style.width = '50px';
        thead.appendChild(thMove);
        savedData.headers.forEach((header, colIdx) => {
            const th = document.createElement('th');
            th.textContent = header.name;
            th.dataset.type = header.type;

            // --- Botão copiar coluna para o próximo mês ---
            const copyBtn = document.createElement('button');
            copyBtn.textContent = '⭳';
            copyBtn.title = 'Copiar esta coluna para o próximo mês';
            copyBtn.style.marginLeft = '5px';
            copyBtn.className = 'copy-col-next-month';
            copyBtn.dataset.colIndex = colIdx;
            copyBtn.dataset.table = 'receitas';
            th.appendChild(copyBtn);

            thead.appendChild(th);
        });

        // Populate rows
        tbody.innerHTML = '';
        savedData.rows.forEach((rowData, idx) => {
            const tr = document.createElement('tr');
            // Adiciona botões de mover linha
            tr.appendChild(createMoveButtons(
                tr,
                (row) => moveRowReceitas(row, -1),
                (row) => moveRowReceitas(row, 1)
            ));
            rowData.forEach((cellData, colIdx) => {
                const td = document.createElement('td');
                // Formata se for coluna de valor
                const header = savedData.headers[colIdx];
                if (header && header.type === 'value' && cellData !== '') {
                    td.textContent = formatCurrency(cellData);
                    td.dataset.rawValue = cellData;
                } else {
                    td.textContent = cellData;
                }
                tr.appendChild(td);
            });
            tbody.appendChild(tr);
        });

        updateSumResult();
    };

    const saveTableData = () => {
        const headers = Array.from(dataTable.querySelector('thead tr').children)
            .slice(1)
            .map(th => ({
                name: th.textContent,
                type: th.dataset.type || 'text'
            }));
        const rows = Array.from(dataTable.querySelector('tbody').children).map(tr =>
            Array.from(tr.children).slice(1).map((td, idx) => {
                // Se for coluna de valor, salva o valor bruto
                const header = headers[idx];
                if (header && header.type === 'value' && td.dataset.rawValue !== undefined) {
                    console.log(header);
                    return td.dataset.rawValue;
                }
                return td.textContent;
            })
        );
        localStorage.setItem(getStorageKeyReceitas(), JSON.stringify({ headers, rows }));
        updateSumResult();
    };

    addRowButton.addEventListener('click', () => {
        const tbody = dataTable.querySelector('tbody');
        const newRow = document.createElement('tr');
        // Adiciona botões de mover linha
        newRow.appendChild(createMoveButtons(
            newRow,
            (row) => moveRowReceitas(row, -1),
            (row) => moveRowReceitas(row, 1)
        ));
        const columnCount = dataTable.querySelector('thead tr').children.length - 1; // ignora coluna dos botões

        // Add empty cells for each column
        for (let i = 0; i < columnCount; i++) {
            const newCell = document.createElement('td');
            newCell.textContent = '';
            newRow.appendChild(newCell);
        }

        tbody.appendChild(newRow);
        saveTableData();
    });

    addColumnButton.addEventListener('click', () => {
        const colName = document.getElementById('colName').value;
        const colType = document.getElementById('colType').value; // Get column type
        const thead = dataTable.querySelector('thead');
        const tbody = dataTable.querySelector('tbody');

        // Add column to the header
        const headerRow = thead.querySelector('tr');
        const newHeader = document.createElement('th');
        newHeader.textContent = colName || `Column ${headerRow.children.length}`;
        newHeader.dataset.type = colType || 'text'; // Set column type
        headerRow.appendChild(newHeader);

        // Add empty cells to existing rows
        tbody.querySelectorAll('tr').forEach(row => {
            const newCell = document.createElement('td');
            newCell.textContent = '';
            row.appendChild(newCell);
        });

        saveTableData();
    });

    let selectedRows = new Set();
    let selectedColumns = new Set();

    // Add click event to rows and columns for multi-selection
    dataTable.addEventListener('click', (event) => {
        const target = event.target;

        if (target.tagName === 'TD') {
            const row = target.parentElement;
            const columnIndex = Array.from(row.children).indexOf(target);

            if (event.ctrlKey) {
                // Toggle row selection
                if (selectedRows.has(row)) {
                    selectedRows.delete(row);
                    row.classList.remove('selected');
                } else {
                    selectedRows.add(row);
                    row.classList.add('selected');
                }
            } else {
                // Clear previous selections
                selectedRows.forEach(r => r.classList.remove('selected'));
                selectedColumns.forEach(index => {
                    dataTable.querySelectorAll('tr').forEach(row => {
                        const cell = row.children[index];
                        if (cell) cell.classList.remove('selected');
                    });
                });
                selectedRows.clear();
                selectedColumns.clear();

                // Select the clicked row
                selectedRows.add(row);
                row.classList.add('selected');
            }
        } else if (target.tagName === 'TH') {
            const columnIndex = Array.from(target.parentElement.children).indexOf(target);

            if (event.ctrlKey) {
                // Toggle column selection
                if (selectedColumns.has(columnIndex)) {
                    selectedColumns.delete(columnIndex);
                    dataTable.querySelectorAll('tr').forEach(row => {
                        const cell = row.children[columnIndex];
                        if (cell) cell.classList.remove('selected');
                    });
                } else {
                    selectedColumns.add(columnIndex);
                    dataTable.querySelectorAll('tr').forEach(row => {
                        const cell = row.children[columnIndex];
                        if (cell) cell.classList.add('selected');
                    });
                }
            } else {
                // Clear previous selections
                selectedRows.forEach(r => r.classList.remove('selected'));
                selectedColumns.forEach(index => {
                    dataTable.querySelectorAll('tr').forEach(row => {
                        const cell = row.children[index];
                        if (cell) cell.classList.remove('selected');
                    });
                });
                selectedRows.clear();
                selectedColumns.clear();

                // Select the clicked column
                selectedColumns.add(columnIndex);
                dataTable.querySelectorAll('tr').forEach(row => {
                    const cell = row.children[columnIndex];
                    if (cell) cell.classList.add('selected');
                });
            }
        } else {
            // Clear all selections if clicking outside
            selectedRows.forEach(r => r.classList.remove('selected'));
            selectedColumns.forEach(index => {
                dataTable.querySelectorAll('tr').forEach(row => {
                    const cell = row.children[index];
                    if (cell) cell.classList.remove('selected');
                });
            });
            selectedRows.clear();
            selectedColumns.clear();
        }

        deleteSelectionButton.style.display = selectedRows.size > 0 || selectedColumns.size > 0 ? 'block' : 'none';
    });

    const deleteSelectionButton = document.createElement('button');
    deleteSelectionButton.id = 'deleteSelection';
    deleteSelectionButton.textContent = 'Excluir Seleção';
    deleteSelectionButton.style.display = 'none'; // Initially hidden
    document.querySelector('.container').appendChild(deleteSelectionButton);

    deleteSelectionButton.addEventListener('click', () => {
        // Delete all selected rows
        selectedRows.forEach(row => row.remove());
        selectedRows.clear();

        // Delete all selected columns
        Array.from(selectedColumns).sort((a, b) => b - a).forEach(index => {
            const thead = dataTable.querySelector('thead tr');
            const tbody = dataTable.querySelector('tbody');

            // Remove header
            if (thead.children[index]) thead.children[index].remove();

            // Remove cells in the column
            tbody.querySelectorAll('tr').forEach(row => {
                if (row.children[index]) row.children[index].remove();
            });
        });
        selectedColumns.clear();

        deleteSelectionButton.style.display = 'none'; // Hide button after deletion
        saveTableData();
    });

    // Make column headers editable
    dataTable.addEventListener('dblclick', (event) => {
        const target = event.target;

        if (target.tagName === 'TH') {
            const currentText = target.textContent;
            const input = document.createElement('input');
            input.type = 'text';
            input.value = currentText;
            input.style.width = '100%';

            // Replace header text with input field
            target.textContent = '';
            target.appendChild(input);

            // Save changes on blur or Enter key
            input.addEventListener('blur', () => {
                target.textContent = input.value || currentText; // Revert if empty
                saveTableData();
            });

            input.addEventListener('keydown', (e) => {
                if (e.key === 'Enter') {
                    input.blur();
                }
            });

            input.focus();
        }
    });

    // Make row cells editable
    dataTable.addEventListener('dblclick', (event) => {
        const target = event.target;

        if (target.tagName === 'TD') {
            // Descobre se é coluna de valor
            const colIdx = Array.from(target.parentElement.children).indexOf(target) - 1;
            const header = dataTable.querySelector('thead tr').children[colIdx + 1];
            const isValue = header && header.dataset.type === 'value';

            // Se for valor, pega valor bruto
            const currentText = isValue && target.dataset.rawValue !== undefined ? target.dataset.rawValue : target.textContent;
            const input = document.createElement('input');
            input.type = 'text';
            input.value = currentText;
            input.style.width = '100%';

            // Replace cell text with input field
            target.textContent = '';
            target.appendChild(input);

            // Save changes on blur or Enter key
            input.addEventListener('blur', () => {
                let value = input.value || currentText;
                if (isValue && value !== '') {
                    target.textContent = formatCurrency(value);
                    target.dataset.rawValue = value;
                } else {
                    target.textContent = value;
                    delete target.dataset.rawValue;
                }
                saveTableData();
                updateSumResult();
            });

            input.addEventListener('keydown', (e) => {
                if (e.key === 'Enter') {
                    input.blur();
                }
            });

            input.focus();
        }
    });

    // Clear selection on Escape key press
    document.addEventListener('keydown', (event) => {
        if (event.key === 'Escape') {
            // Clear row and column selections
            dataTable.querySelectorAll('tr').forEach(row => row.classList.remove('selected'));
            dataTable.querySelectorAll('th, td').forEach(cell => cell.classList.remove('selected'));
            selectedRows.clear();
            selectedColumns.clear();
            deleteSelectionButton.style.display = 'none'; // Hide button
        }
    });

    // Function to calculate and display the sum of "Valor" columns
    const updateSumResult = () => {
        const tbody = dataTable.querySelector('tbody');
        const headers = Array.from(dataTable.querySelector('thead tr').children);
        const valueColumnIndexes = headers
            .map((th, index) => (th.dataset.type === 'value' ? index : -1))
            .filter(index => index !== -1);

        let sum = 0;
        tbody.querySelectorAll('tr').forEach(row => {
            valueColumnIndexes.forEach(index => {
                const cell = row.children[index];
                if (cell) {
                    // Sempre usa valor bruto se disponível
                    const value = parseFloat(cell.dataset.rawValue !== undefined ? cell.dataset.rawValue : cell.textContent.replace(/[^\d\.,-]/g, '').replace(',', '.')) || 0;
                    sum += value;
                }
            });
        });

        document.getElementById('sumResult').textContent = formatCurrency(sum);
    };

    // Load table data on page load
    loadTableData();

    const addRowButtonExpenses = document.getElementById('addRowExpenses');
    const addColumnButtonExpenses = document.getElementById('addColumnExpenses');
    const dataTableExpenses = document.getElementById('dataTableExpenses');

    const loadTableDataExpenses = () => {
        const savedData = JSON.parse(localStorage.getItem(getStorageKeyDespesas())) || { headers: [], rows: [] };
        const thead = dataTableExpenses.querySelector('thead tr');
        const tbody = dataTableExpenses.querySelector('tbody');

        // Populate headers
        thead.innerHTML = '';
        // Header para botões
        const thMove = document.createElement('th');
        thMove.style.width = '50px';
        thead.appendChild(thMove);
        savedData.headers.forEach((header, colIdx) => {
            const th = document.createElement('th');
            th.textContent = header.name;
            th.dataset.type = header.type;

            // --- Botão copiar coluna para o próximo mês ---
            const copyBtn = document.createElement('button');
            copyBtn.textContent = '⭳';
            copyBtn.title = 'Copiar esta coluna para o próximo mês';
            copyBtn.style.marginLeft = '5px';
            copyBtn.className = 'copy-col-next-month';
            copyBtn.dataset.colIndex = colIdx;
            copyBtn.dataset.table = 'despesas';
            th.appendChild(copyBtn);

            thead.appendChild(th);
        });

        // Populate rows
        tbody.innerHTML = '';
        savedData.rows.forEach((rowData, idx) => {
            const tr = document.createElement('tr');
            tr.appendChild(createMoveButtons(
                tr,
                (row) => moveRowDespesas(row, -1),
                (row) => moveRowDespesas(row, 1)
            ));
            rowData.forEach((cellData, colIdx) => {
                const td = document.createElement('td');
                const header = savedData.headers[colIdx];
                if (header && header.type === 'value' && cellData !== '') {
                    td.textContent = formatCurrency(cellData);
                    td.dataset.rawValue = cellData;
                } else {
                    td.textContent = cellData;
                }
                tr.appendChild(td);
            });
            tbody.appendChild(tr);
        });

        updateSumResultExpenses();
    };

    const saveTableDataExpenses = () => {
        const headers = Array.from(dataTableExpenses.querySelector('thead tr').children)
            .slice(1)
            .map(th => ({
                name: th.textContent,
                type: th.dataset.type || 'text'
            }));
        const rows = Array.from(dataTableExpenses.querySelector('tbody').children).map(tr =>
            Array.from(tr.children).slice(1).map((td, idx) => {
                const header = headers[idx];
                if (header && header.type === 'value' && td.dataset.rawValue !== undefined) {
                    return td.dataset.rawValue;
                }
                return td.textContent;
            })
        );
        localStorage.setItem(getStorageKeyDespesas(), JSON.stringify({ headers, rows }));
        updateSumResultExpenses();
    };

    addRowButtonExpenses.addEventListener('click', () => {
        const tbody = dataTableExpenses.querySelector('tbody');
        const newRow = document.createElement('tr');
        newRow.appendChild(createMoveButtons(
            newRow,
            (row) => moveRowDespesas(row, -1),
            (row) => moveRowDespesas(row, 1)
        ));
        const columnCount = dataTableExpenses.querySelector('thead tr').children.length - 1;

        for (let i = 0; i < columnCount; i++) {
            const newCell = document.createElement('td');
            newCell.textContent = '';
            newRow.appendChild(newCell);
        }

        tbody.appendChild(newRow);
        saveTableDataExpenses();
    });

    addColumnButtonExpenses.addEventListener('click', () => {
        const colName = document.getElementById('colNameExpenses').value;
        const colType = document.getElementById('colTypeExpenses').value;
        const thead = dataTableExpenses.querySelector('thead');
        const tbody = dataTableExpenses.querySelector('tbody');

        const headerRow = thead.querySelector('tr');
        const newHeader = document.createElement('th');
        newHeader.textContent = colName || `Column ${headerRow.children.length}`;
        newHeader.dataset.type = colType || 'text';
        headerRow.appendChild(newHeader);

        tbody.querySelectorAll('tr').forEach(row => {
            const newCell = document.createElement('td');
            newCell.textContent = '';
            row.appendChild(newCell);
        });

        saveTableDataExpenses();
    });

    const updateSumResultExpenses = () => {
        const tbody = dataTableExpenses.querySelector('tbody');
        const headers = Array.from(dataTableExpenses.querySelector('thead tr').children);
        const valueColumnIndexes = headers
            .map((th, index) => (th.dataset.type === 'value' ? index : -1))
            .filter(index => index !== -1);

        let sum = 0;
        tbody.querySelectorAll('tr').forEach(row => {
            valueColumnIndexes.forEach(index => {
                const cell = row.children[index];
                if (cell) {
                    const value = parseFloat(cell.dataset.rawValue !== undefined ? cell.dataset.rawValue : cell.textContent.replace(/[^\d\.,-]/g, '').replace(',', '.')) || 0;
                    sum += value;
                }
            });
        });

        document.getElementById('sumResultExpenses').textContent = formatCurrency(sum);
    };

    // Load table data for Expenses on page load
    loadTableDataExpenses();

    // Add similar event listeners for Expenses table (e.g., row/column selection, editing, etc.)
    // Make column headers editable for Expenses
    dataTableExpenses.addEventListener('dblclick', (event) => {
        const target = event.target;

        if (target.tagName === 'TH') {
            const currentText = target.textContent;
            const input = document.createElement('input');
            input.type = 'text';
            input.value = currentText;
            input.style.width = '100%';

            // Replace header text with input field
            target.textContent = '';
            target.appendChild(input);

            // Save changes on blur or Enter key
            input.addEventListener('blur', () => {
                target.textContent = input.value || currentText; // Revert if empty
                saveTableDataExpenses();
            });

            input.addEventListener('keydown', (e) => {
                if (e.key === 'Enter') {
                    input.blur();
                }
            });

            input.focus();
        }
    });

    // Make row cells editable for Expenses
    dataTableExpenses.addEventListener('dblclick', (event) => {
        const target = event.target;

        if (target.tagName === 'TD') {
            const colIdx = Array.from(target.parentElement.children).indexOf(target) - 1;
            const header = dataTableExpenses.querySelector('thead tr').children[colIdx + 1];
            const isValue = header && header.dataset.type === 'value';

            const currentText = isValue && target.dataset.rawValue !== undefined ? target.dataset.rawValue : target.textContent;
            const input = document.createElement('input');
            input.type = 'text';
            input.value = currentText;
            input.style.width = '100%';

            // Replace cell text with input field
            target.textContent = '';
            target.appendChild(input);

            // Save changes on blur or Enter key
            input.addEventListener('blur', () => {
                let value = input.value || currentText;
                if (isValue && value !== '') {
                    target.textContent = formatCurrency(value);
                    target.dataset.rawValue = value;
                } else {
                    target.textContent = value;
                    delete target.dataset.rawValue;
                }
                saveTableDataExpenses();
                updateSumResultExpenses();
            });

            input.addEventListener('keydown', (e) => {
                if (e.key === 'Enter') {
                    input.blur();
                }
            });

            input.focus();
        }
    });

    

    // Clear selection on Escape key press for Expenses
    document.addEventListener('keydown', (event) => {
        if (event.key === 'Escape') {
            // Clear row and column selections for Expenses
            dataTableExpenses.querySelectorAll('tr').forEach(row => row.classList.remove('selected'));
            dataTableExpenses.querySelectorAll('th, td').forEach(cell => cell.classList.remove('selected'));
            // ...existing code for clearing selections...
        }
    });

    // --- NOVAS FUNÇÕES PARA RESULTADO FINAL, LIMPAR, EXPORTAR, IMPORTAR, TEMPLATE ---

    // Calcular Resultado Final (Receitas - Despesas)
    document.getElementById('calculateFinalResult').addEventListener('click', () => {
        const receitas = parseCurrencyText(document.getElementById('sumResult').textContent);
        const despesas = parseCurrencyText(document.getElementById('sumResultExpenses').textContent);
        const poupanca = parseCurrencyText(document.getElementById('sumResultSavings').textContent);
        const resultadoFinal = receitas - (despesas - poupanca);
        document.getElementById('finalResult').textContent = resultadoFinal.toFixed(2);
    });

    // Calcular Resultado Final (Receitas - Despesas - Poupança do mês atual)
    document.getElementById('calculateFinalResult').addEventListener('click', () => {
        const receitas = parseCurrencyText(document.getElementById('sumResult').textContent);
        const despesas = parseCurrencyText(document.getElementById('sumResultExpenses').textContent);
        const poupancaMesAtual = parseCurrencyText(document.getElementById('sumResultSavings').textContent);
        const resultadoFinal = receitas - despesas - poupancaMesAtual;
        document.getElementById('finalResult').textContent = resultadoFinal.toFixed(2);
    });

    // Calcular Resultado Final (Receitas - Despesas - Poupança do mês atual - Reserva de Emergência do mês atual)
    document.getElementById('calculateFinalResult').addEventListener('click', () => {
        const receitas = parseCurrencyText(document.getElementById('sumResult').textContent);
        const despesas = parseCurrencyText(document.getElementById('sumResultExpenses').textContent);
        const poupancaMesAtual = parseCurrencyText(document.getElementById('sumResultSavings').textContent);
        const emergenciaMesAtual = parseCurrencyText(document.getElementById('sumResultEmergency').textContent);
        const resultadoFinal = receitas - despesas - poupancaMesAtual - emergenciaMesAtual;
        document.getElementById('finalResult').textContent = resultadoFinal.toFixed(2);
    });

    // Limpar Tudo
    document.getElementById('clearAll').addEventListener('click', () => {
        if (confirm('Tem certeza que deseja limpar todos os dados? Esta ação não pode ser desfeita.')) {
            localStorage.clear();
            // Atualiza a interface após limpar
            document.getElementById('finalResult').textContent = '0';
            // Recarrega tabelas e campos
            if (typeof loadAllMonthData === 'function') {
                loadAllMonthData();
            } else {
                // fallback para funções individuais
                loadTableData && loadTableData();
                loadTableDataExpenses && loadTableDataExpenses();
                loadSavingsValue && loadSavingsValue();
                loadNextMonthValue && loadNextMonthValue();
            }
            // Limpa inputs
            document.getElementById('savingsValue').value = '';
            document.getElementById('emergencyValue').value = '';
            document.getElementById('nextMonthValue').value = '';
            document.getElementById('sumResult').textContent = '€0.00';
            document.getElementById('sumResultExpenses').textContent = '€0.00';
            document.getElementById('sumResultSavings').textContent = '€0.00';
            document.getElementById('sumResultSavingsTotal').textContent = '€0.00';
            document.getElementById('sumResultEmergency').textContent = '€0.00';
            document.getElementById('sumResultEmergencyTotal').textContent = '€0.00';
        }
    });

    // Exportar Dados (agora em formato Excel)
    document.getElementById('exportData').addEventListener('click', () => {
        const wb = XLSX.utils.book_new();

        // Para cada mês do ano atual (ou todos os anos, se desejar)
        // Aqui, exporta apenas o ano atual
        for (let year = currentYear; year <= currentYear; year++) {
            for (let month = 1; month <= 12; month++) {
                // Receitas
                const receitasData = JSON.parse(localStorage.getItem(`tableData_${year}_${month}`)) || { headers: [], rows: [] };
                const receitasSheet = [
                    receitasData.headers.map(h => h.name),
                    ...receitasData.rows
                ];

                // Despesas
                const despesasData = JSON.parse(localStorage.getItem(`tableDataExpenses_${year}_${month}`)) || { headers: [], rows: [] };
                const despesasSheet = [
                    despesasData.headers.map(h => h.name),
                    ...despesasData.rows
                ];

                // Poupança e Emergência
                const savingsValue = parseFloat(localStorage.getItem(`savingsValue_${year}_${month}`)) || 0;
                const emergencyValue = parseFloat(localStorage.getItem(`emergencyValue_${year}_${month}`)) || 0;

                // Monta sheet: Receitas, linha em branco, Despesas, linha em branco, Poupança/Emergência
                let sheetData = [];
                if (receitasSheet.length > 1) {
                    sheetData.push(['Receitas']);
                    sheetData = sheetData.concat(receitasSheet);
                    sheetData.push([]);
                }
                if (despesasSheet.length > 1) {
                    sheetData.push(['Despesas']);
                    sheetData = sheetData.concat(despesasSheet);
                    sheetData.push([]);
                }
                sheetData.push(['Poupança', savingsValue]);
                sheetData.push(['Reserva de Emergência', emergencyValue]);

                // Nome da aba: MM-YYYY
                const sheetName = `${String(month).padStart(2, '0')}-${year}`;
                const ws = XLSX.utils.aoa_to_sheet(sheetData);
                XLSX.utils.book_append_sheet(wb, ws, sheetName);
            }
        }

        XLSX.writeFile(wb, `dados-financeiros-${currentYear}.xlsx`);
    });

    // Importar Dados (agora aceita .xlsx exportado)
    document.getElementById('importData').addEventListener('click', () => {
        const fileInput = document.getElementById('fileInput');
        fileInput.accept = '.xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
        fileInput.value = ''; // Limpa para permitir novo upload igual
        fileInput.click();
    });

    document.getElementById('fileInput').addEventListener('change', (event) => {
        const file = event.target.files[0];
        if (!file) return;

        if (file.name.endsWith('.xlsx')) {
            const reader = new FileReader();
            reader.onload = (e) => {
                try {
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, { type: 'array' });

                    // Para cada sheet (mês)
                    workbook.SheetNames.forEach(sheetName => {
                        // Esperado: sheetName = MM-YYYY
                        const [monthStr, yearStr] = sheetName.split('-');
                        const month = parseInt(monthStr, 10);
                        const year = parseInt(yearStr, 10);
                        const aoa = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1 });

                        // Procura blocos Receitas, Despesas, Poupança, Emergência
                        let receitas = { headers: [], rows: [] };
                        let despesas = { headers: [], rows: [] };
                        let savingsValue = 0;
                        let emergencyValue = 0;

                        let i = 0;
                        while (i < aoa.length) {
                            const row = aoa[i];
                            if (row[0] === 'Receitas') {
                                i++;
                                receitas.headers = (aoa[i] || []).map(name => ({
                                    name,
                                    type: name && name.toLowerCase().includes('valor') ? 'value' : 'text'
                                }));
                                i++;
                                while (i < aoa.length && aoa[i].length && aoa[i][0] !== 'Despesas') {
                                    receitas.rows.push(aoa[i]);
                                    i++;
                                }
                            } else if (row[0] === 'Despesas') {
                                i++;
                                despesas.headers = (aoa[i] || []).map(name => ({
                                    name,
                                    type: name && name.toLowerCase().includes('valor') ? 'value' : 'text'
                                }));
                                i++;
                                while (i < aoa.length && aoa[i].length && aoa[i][0] !== 'Poupança' && aoa[i][0] !== 'Reserva de Emergência') {
                                    despesas.rows.push(aoa[i]);
                                    i++;
                                }
                            } else if (row[0] === 'Poupança') {
                                savingsValue = parseFloat(row[1]) || 0;
                                i++;
                            } else if (row[0] === 'Reserva de Emergência') {
                                emergencyValue = parseFloat(row[1]) || 0;
                                i++;
                            } else {
                                i++;
                            }
                        }

                        // Salva no localStorage
                        if (year && month) {
                            localStorage.setItem(`tableData_${year}_${month}`, JSON.stringify(receitas));
                            localStorage.setItem(`tableDataExpenses_${year}_${month}`, JSON.stringify(despesas));
                            localStorage.setItem(`savingsValue_${year}_${month}`, savingsValue);
                            localStorage.setItem(`emergencyValue_${year}_${month}`, emergencyValue);
                        }
                    });

                    // Após importar, recarrega dados do mês atual
                    updateMonthLabel();
                    loadTableData();
                    loadTableDataExpenses();
                    loadSavingsValue();
                    loadNextMonthValue && loadNextMonthValue();
                    updateSumResultSavingsTotal();
                    document.getElementById('finalResult').textContent = '0';
                } catch {
                    alert('Erro ao importar arquivo Excel.');
                }
            };
            reader.readAsArrayBuffer(file);
        } else {
            alert('Por favor, selecione um arquivo .xlsx exportado pelo sistema.');
        }
        event.target.value = '';
    });

  

    // --- Seleção e exclusão para DESPESAS ---
    let selectedRowsExpenses = new Set();
    let selectedColumnsExpenses = new Set();

    dataTableExpenses.addEventListener('click', (event) => {
        const target = event.target;

        if (target.tagName === 'TD') {
            const row = target.parentElement;
            const columnIndex = Array.from(row.children).indexOf(target);

            if (event.ctrlKey) {
                // Toggle row selection
                if (selectedRowsExpenses.has(row)) {
                    selectedRowsExpenses.delete(row);
                    row.classList.remove('selected');
                } else {
                    selectedRowsExpenses.add(row);
                    row.classList.add('selected');
                }
            } else {
                // Clear previous selections
                selectedRowsExpenses.forEach(r => r.classList.remove('selected'));
                selectedColumnsExpenses.forEach(index => {
                    dataTableExpenses.querySelectorAll('tr').forEach(row => {
                        const cell = row.children[index];
                        if (cell) cell.classList.remove('selected');
                    });
                });
                selectedRowsExpenses.clear();
                selectedColumnsExpenses.clear();

                // Select the clicked row
                selectedRowsExpenses.add(row);
                row.classList.add('selected');
            }
        } else if (target.tagName === 'TH') {
            const columnIndex = Array.from(target.parentElement.children).indexOf(target);

            if (event.ctrlKey) {
                // Toggle column selection
                if (selectedColumnsExpenses.has(columnIndex)) {
                    selectedColumnsExpenses.delete(columnIndex);
                    dataTableExpenses.querySelectorAll('tr').forEach(row => {
                        const cell = row.children[columnIndex];
                        if (cell) cell.classList.remove('selected');
                    });
                } else {
                    selectedColumnsExpenses.add(columnIndex);
                    dataTableExpenses.querySelectorAll('tr').forEach(row => {
                        const cell = row.children[columnIndex];
                        if (cell) cell.classList.add('selected');
                    });
                }
            } else {
                // Clear previous selections
                selectedRowsExpenses.forEach(r => r.classList.remove('selected'));
                selectedColumnsExpenses.forEach(index => {
                    dataTableExpenses.querySelectorAll('tr').forEach(row => {
                        const cell = row.children[index];
                        if (cell) cell.classList.remove('selected');
                    });
                });
                selectedRowsExpenses.clear();
                selectedColumnsExpenses.clear();

                // Select the clicked column
                selectedColumnsExpenses.add(columnIndex);
                dataTableExpenses.querySelectorAll('tr').forEach(row => {
                    const cell = row.children[columnIndex];
                    if (cell) cell.classList.add('selected');
                });
            }
        } else {
            // Clear all selections if clicking outside
            selectedRowsExpenses.forEach(r => r.classList.remove('selected'));
            selectedColumnsExpenses.forEach(index => {
                dataTableExpenses.querySelectorAll('tr').forEach(row => {
                    const cell = row.children[index];
                    if (cell) cell.classList.remove('selected');
                });
            });
            selectedRowsExpenses.clear();
            selectedColumnsExpenses.clear();
        }

        deleteSelectionButtonExpenses.style.display = selectedRowsExpenses.size > 0 || selectedColumnsExpenses.size > 0 ? 'block' : 'none';
    });

    const deleteSelectionButtonExpenses = document.createElement('button');
    deleteSelectionButtonExpenses.id = 'deleteSelectionExpenses';
    deleteSelectionButtonExpenses.textContent = 'Excluir Seleção';
    deleteSelectionButtonExpenses.style.display = 'none'; // Initially hidden
    // Adiciona o botão na segunda .container (Despesas)
    document.querySelectorAll('.container')[1].appendChild(deleteSelectionButtonExpenses);

    deleteSelectionButtonExpenses.addEventListener('click', () => {
        // Delete all selected rows
        selectedRowsExpenses.forEach(row => row.remove());
        selectedRowsExpenses.clear();

        // Delete all selected columns
        Array.from(selectedColumnsExpenses).sort((a, b) => b - a).forEach(index => {
            const thead = dataTableExpenses.querySelector('thead tr');
            const tbody = dataTableExpenses.querySelector('tbody');

            // Remove header
            if (thead.children[index]) thead.children[index].remove();

            // Remove cells in the column
            tbody.querySelectorAll('tr').forEach(row => {
                if (row.children[index]) row.children[index].remove();
            });
        });
        selectedColumnsExpenses.clear();

        deleteSelectionButtonExpenses.style.display = 'none'; // Hide button after deletion
        saveTableDataExpenses();
    });

    // Clear selection on Escape key press for Expenses
    document.addEventListener('keydown', (event) => {
        if (event.key === 'Escape') {
            // Clear row and column selections for Expenses
            dataTableExpenses.querySelectorAll('tr').forEach(row => row.classList.remove('selected'));
            dataTableExpenses.querySelectorAll('th, td').forEach(cell => cell.classList.remove('selected'));
            selectedRowsExpenses.clear();
            selectedColumnsExpenses.clear();
            deleteSelectionButtonExpenses.style.display = 'none';
            // ...existing code for clearing selections...
        }
    });

    // --- FIM DAS ALTERAÇÕES DE PAGINAÇÃO POR MÊS ---

    // Carrega dados do mês atual ao iniciar
    loadTableData();
    loadTableDataExpenses();

    // --- NAVEGAÇÃO ENTRE MESES ---
    prevBtn.addEventListener('click', () => {
        if (currentMonth === 0) {
            currentMonth = 11;
            currentYear--;
        } else {
            currentMonth;
        }
        updateMonthLabel();
        loadTableData();
        loadTableDataExpenses();
        document.getElementById('finalResult').textContent = '0';
    });

    nextBtn.addEventListener('click', () => {
        if (currentMonth === 11) {
            currentMonth = 0;
            currentYear++;
        } else {
            currentMonth;
        }
        updateMonthLabel();
        loadTableData();
        loadTableDataExpenses();
        document.getElementById('finalResult').textContent = '0';
    });

    // Atualiza label e carrega dados do mês inicial
    updateMonthLabel();

    // ...restante do código permanece igual...

    // --- POUPANÇA (apenas valor) ---

    function getStorageKeySavings(month = null, year = null) {
        if (month === null) month = currentMonth + 1;
        if (year === null) year = currentYear;
        return `savingsValue_${year}_${month}`;
    }
    function getStorageKeyEmergency(month = null, year = null) {
        if (month === null) month = currentMonth + 1;
        if (year === null) year = currentYear;
        return `emergencyValue_${year}_${month}`;
    }

    const savingsInput = document.getElementById('savingsValue');
    const sumResultSavings = document.getElementById('sumResultSavings');
    const sumResultSavingsTotal = document.getElementById('sumResultSavingsTotal');

    // NOVO: Reserva de Emergência
    const emergencyInput = document.getElementById('emergencyValue');
    const sumResultEmergency = document.getElementById('sumResultEmergency');
    const sumResultEmergencyTotal = document.getElementById('sumResultEmergencyTotal');

    // Defina a função antes de usá-la!
    function loadEmergencyValue() {
        const value = parseFloat(localStorage.getItem(getStorageKeyEmergency())) || 0;
        emergencyInput.value = value !== 0 ? value : '';
        sumResultEmergency.textContent = formatCurrency(value);
        updateSumResultEmergencyTotal();
    }

    function saveEmergencyValue() {
        const value = parseFloat(emergencyInput.value) || 0;
        localStorage.setItem(getStorageKeyEmergency(), value);
        sumResultEmergency.textContent = formatCurrency(value);
        updateSumResultEmergencyTotal();
    }

    function updateSumResultEmergencyTotal() {
        let total = 0;
        // Soma apenas meses anteriores e o mês atual do ano corrente
        for (let m = 1; m <= currentMonth + 1; m++) {
            if (m === currentMonth + 1) {
                const val = parseFloat(emergencyInput.value) || 0;
                total += val;
            } else {
                const val = parseFloat(localStorage.getItem(`emergencyValue_${currentYear}_${m}`)) || 0;
                total += val;
            }
        }
        sumResultEmergencyTotal.textContent = formatCurrency(total);
    }

    function loadSavingsValue() {
        const value = parseFloat(localStorage.getItem(getStorageKeySavings())) || 0;
        savingsInput.value = value !== 0 ? value : '';
        sumResultSavings.textContent = formatCurrency(value);
        updateSumResultSavingsTotal();
        loadEmergencyValue(); // agora a função está definida acima
    }

    function saveSavingsValue() {
        const value = parseFloat(savingsInput.value) || 0;
        localStorage.setItem(getStorageKeySavings(), value);
        sumResultSavings.textContent = formatCurrency(value);
        updateSumResultSavingsTotal();
    }

    function saveEmergencyValue() {
        const value = parseFloat(emergencyInput.value) || 0;
        localStorage.setItem(getStorageKeyEmergency(), value);
        sumResultEmergency.textContent = formatCurrency(value);
        updateSumResultEmergencyTotal();
    }

    function updateSumResultSavingsTotal() {
        let total = 0;
        // Soma apenas meses anteriores e o mês atual do ano corrente
        for (let m = 1; m <= currentMonth + 1; m++) {
            if (m === currentMonth + 1) {
                // mês atual: pega do input (pode estar alterado e não salvo ainda)
                const val = parseFloat(savingsInput.value) || 0;
                total += val;
            } else {
                const val = parseFloat(localStorage.getItem(`savingsValue_${currentYear}_${m}`)) || 0;
                total += val;
            }
        }
        sumResultSavingsTotal.textContent = formatCurrency(total);
    }

    // Atualize o total acumulado ao digitar, mesmo sem salvar
    savingsInput.addEventListener('input', () => {
        saveSavingsValue();
        updateSumResultSavingsTotal();
    });
    emergencyInput.addEventListener('input', () => {
        saveEmergencyValue();
        updateSumResultEmergencyTotal();
    });

    // --- NOVAS FUNÇÕES PARA RESULTADO FINAL, LIMPAR, EXPORTAR, IMPORTAR, TEMPLATE ---

    // Calcular Resultado Final (Receitas - Despesas - Poupança acumulada)
    document.getElementById('calculateFinalResult').addEventListener('click', () => {
        const receitas = parseCurrencyText(document.getElementById('sumResult').textContent);
        const despesas = parseCurrencyText(document.getElementById('sumResultExpenses').textContent);
        const poupancaTotal = parseCurrencyText(sumResultSavingsTotal.textContent);
        const emergenciaMesAtual = parseCurrencyText(document.getElementById('sumResultEmergency').textContent);
        const nextMonthVal = parseFloat(document.getElementById('nextMonthValue').value) || 0;
        const resultadoFinal = receitas - despesas - poupancaTotal - emergenciaMesAtual - nextMonthVal;
        document.getElementById('finalResult').textContent = formatCurrency(resultadoFinal);
    });

    // Remova os outros listeners duplicados para o botão 'calculateFinalResult'
    // ...existing code...

    // Garanta que updateSumResultSavingsTotal é chamada sempre que savingsInput muda
    savingsInput.addEventListener('input', () => {
        saveSavingsValue();
        updateSumResultSavingsTotal();
    });

    // Garanta que updateSumResultSavingsTotal é chamada ao trocar de mês
    function loadSavingsValue() {
        const value = parseFloat(localStorage.getItem(getStorageKeySavings())) || 0;
        savingsInput.value = value !== 0 ? value : '';
        sumResultSavings.textContent = formatCurrency(value);
        updateSumResultSavingsTotal();
        loadEmergencyValue(); // agora a função está definida acima
    }

    // ...existing code...

    // Garanta que updateSumResultSavingsTotal é chamada ao importar dados
    document.getElementById('fileInput').addEventListener('change', (event) => {
        // ...existing code...
        if (file.name.endsWith('.xlsx')) {
            // ...existing code...
            reader.onload = (e) => {
                try {
                    // ...existing code...
                    // Após importar, recarrega dados do mês atual
                    updateMonthLabel();
                    loadTableData();
                    loadTableDataExpenses();
                    loadSavingsValue();
                    loadNextMonthValue && loadNextMonthValue();
                    updateSumResultSavingsTotal();
                    document.getElementById('finalResult').textContent = '0';
                } catch {
                    alert('Erro ao importar arquivo Excel.');
                }
            };
            // ...existing code...
        }
        // ...existing code...
    });

    // ...existing code...

    // Garanta que updateSumResultSavingsTotal é chamada ao navegar entre meses
    function loadAllMonthData() {
        loadTableData();
        loadTableDataExpenses();
        loadSavingsValue();
        loadNextMonthValue();
        updateSumResultSavingsTotal();
        document.getElementById('finalResult').textContent = '0';
    }

    // --- Próximo Mês ---
    function getStorageKeyNextMonth(month = null, year = null) {
        if (month === null) month = currentMonth + 1;
        if (year === null) year = currentYear;
        return `nextMonthValue_${year}_${month}`;
    }

    const nextMonthInput = document.getElementById('nextMonthValue');

    function loadNextMonthValue() {
        const value = parseFloat(localStorage.getItem(getStorageKeyNextMonth())) || 0;
        nextMonthInput.value = value !== 0 ? value : '';
    }

    function saveNextMonthValue() {
        const value = parseFloat(nextMonthInput.value) || 0;
        localStorage.setItem(getStorageKeyNextMonth(), value);
        // Atualiza saldo no mês seguinte
        insertSaldoAnteriorInNextMonth(value);
    }

    function insertSaldoAnteriorInNextMonth(value) {
        // Calcula próximo mês/ano
        let nextMonth = currentMonth + 2;
        let nextYear = currentYear;
        if (nextMonth > 12) {
            nextMonth = 1;
            nextYear++;
        }
        // Carrega dados do mês seguinte
        const key = `tableData_${nextYear}_${nextMonth}`;
        let data = JSON.parse(localStorage.getItem(key)) || { headers: [], rows: [] };

        // Garante colunas mínimas: Data, Receita, Valor
        let idxData = data.headers.findIndex(h => h.name === 'Data');
        let idxReceita = data.headers.findIndex(h => h.name === 'Receita');
        let idxValor = data.headers.findIndex(h => h.name === 'Valor');
        if (idxData === -1) {
            data.headers.unshift({ name: 'Data', type: 'date' });
            idxData = 0;
            idxReceita++;
            idxValor++;
        }
        if (idxReceita === -1) {
            data.headers.splice(1, 0, { name: 'Receita', type: 'text' });
            idxReceita = 1;
            idxValor++;
        }
        if (idxValor === -1) {
            data.headers.push({ name: 'Valor', type: 'value' });
            idxValor = data.headers.length - 1;
        }

        // Cria linha "Saldo anterior"
        const firstDay = `${String(nextYear).padStart(4, '0')}-${String(nextMonth).padStart(2, '0')}-01`;
        let saldoRow = Array(data.headers.length).fill('');
        saldoRow[idxData] = firstDay;
        saldoRow[idxReceita] = 'Saldo anterior';
        saldoRow[idxValor] = value.toFixed(2);

        // Remove linha "Saldo anterior" existente (se houver)
        data.rows = data.rows.filter(row => !(row[idxReceita] === 'Saldo anterior' && row[idxData] === firstDay));
        // Insere no topo
        data.rows.unshift(saldoRow);

        localStorage.setItem(key, JSON.stringify(data));
    }

    nextMonthInput.addEventListener('input', () => {
        saveNextMonthValue();
        updateFinalResult();
    });

    // Botão para registrar explicitamente o valor no próximo mês
    const registerNextMonthBtn = document.getElementById('registerNextMonth');
    registerNextMonthBtn.addEventListener('click', () => {
        const value = parseFloat(nextMonthInput.value) || 0;
        insertSaldoAnteriorInNextMonth(value);
        // Opcional: feedback visual
        registerNextMonthBtn.textContent = 'Registrado!';
        setTimeout(() => { registerNextMonthBtn.textContent = 'Registrar no próximo mês'; }, 1000);
    });

    function updateFinalResult() {
        const receitas = parseCurrencyText(document.getElementById('sumResult').textContent);
        const despesas = parseCurrencyText(document.getElementById('sumResultExpenses').textContent);
        const poupancaMesAtual = parseCurrencyText(document.getElementById('sumResultSavings').textContent);
        const emergenciaMesAtual = parseCurrencyText(document.getElementById('sumResultEmergency').textContent);
        const nextMonthVal = parseFloat(nextMonthInput.value) || 0;
        const resultadoFinal = receitas - despesas - poupancaMesAtual - emergenciaMesAtual - nextMonthVal;
        document.getElementById('finalResult').textContent = formatCurrency(resultadoFinal);
    }

    // --- Ajuste: ao calcular resultado final, subtrai também o valor do próximo mês ---
    document.getElementById('calculateFinalResult').addEventListener('click', updateFinalResult);

    // --- Carregar valores ao trocar de mês ---
    function loadAllMonthData() {
        loadTableData();
        loadTableDataExpenses();
        loadSavingsValue();
        loadNextMonthValue();
        updateSumResultSavingsTotal();
        document.getElementById('finalResult').textContent = '0';
    }

    // --- NAVEGAÇÃO ENTRE MESES ---
    prevBtn.addEventListener('click', () => {
        if (currentMonth === 0) {
            currentMonth = 11;
            currentYear--;
        } else {
            currentMonth--;
        }
        updateMonthLabel();
        loadAllMonthData();
    });

    nextBtn.addEventListener('click', () => {
        if (currentMonth === 11) {
            currentMonth = 0;
            currentYear++;
        } else {
            currentMonth++;
        }
        updateMonthLabel();
        loadAllMonthData();
    });

    // --- Carrega dados do mês atual ao iniciar ---
    updateMonthLabel();
    loadAllMonthData();

    // --- Ao importar dados, carregar também o valor do próximo mês ---
    // (Ajuste nos listeners de importação, se necessário)
    // ...existing code...

    // --- Handler para copiar coluna para o próximo mês ---
    function copyColumnToNextMonth(tableType, colIndex) {
        // Descobre mês/ano seguinte
        let nextMonth = currentMonth + 2;
        let nextYear = currentYear;
        if (nextMonth > 12) {
            nextMonth = 1;
            nextYear++;
        }
        // Chaves e elementos
        let getStorageKey, dataTableEl;
        if (tableType === 'receitas') {
            getStorageKey = (y, m) => `tableData_${y}_${m}`;
            dataTableEl = dataTable;
        } else {
            getStorageKey = (y, m) => `tableDataExpenses_${y}_${m}`;
            dataTableEl = dataTableExpenses;
        }
        // Dados do mês atual
        const currentData = JSON.parse(localStorage.getItem(getStorageKey(currentYear, currentMonth + 1))) || { headers: [], rows: [] };
        // Dados do mês seguinte
        let nextData = JSON.parse(localStorage.getItem(getStorageKey(nextYear, nextMonth))) || { headers: [], rows: [] };

        // Header a copiar
        const headerToCopy = currentData.headers[colIndex];
        if (!headerToCopy) return;

        // Verifica se já existe coluna igual no mês seguinte
        let targetColIdx = nextData.headers.findIndex(h => h.name === headerToCopy.name);
        if (targetColIdx === -1) {
            // Adiciona coluna ao final
            nextData.headers.push({ ...headerToCopy });
            targetColIdx = nextData.headers.length - 1;
            // Adiciona célula vazia nas linhas existentes
            nextData.rows.forEach(row => row.push(''));
        }

        // Copia valores das linhas (por índice)
        // Garante que nextData.rows tenha pelo menos tantas linhas quanto currentData.rows
        while (nextData.rows.length < currentData.rows.length) {
            // Preenche linhas faltantes com células vazias
            nextData.rows.push(Array(nextData.headers.length).fill(''));
        }
        for (let i = 0; i < currentData.rows.length; i++) {
            // Garante que a linha tenha o tamanho correto
            while (nextData.rows[i].length < nextData.headers.length) {
                nextData.rows[i].push('');
            }
            nextData.rows[i][targetColIdx] = currentData.rows[i][colIndex] || '';
        }

        // Salva no localStorage
        localStorage.setItem(getStorageKey(nextYear, nextMonth), JSON.stringify(nextData));
        // Feedback visual opcional
        alert(`Coluna "${headerToCopy.name}" copiada para o mês ${nextMonth}/${nextYear}.`);
    }

    // --- Delegação de evento para botões de copiar coluna (Receitas e Despesas) ---
    document.body.addEventListener('click', (e) => {
        if (e.target.classList.contains('copy-col-next-month')) {
            const colIdx = parseInt(e.target.dataset.colIndex, 10);
            const tableType = e.target.dataset.table;
            copyColumnToNextMonth(tableType, colIdx);
        }
    });

    // --- Função utilitária para criar botões de mover linha ---
    function createMoveButtons(row, moveUpCallback, moveDownCallback) {
        // Cria célula para botões
        const td = document.createElement('td');
        td.style.whiteSpace = 'nowrap';
        // Botão para cima
        const upBtn = document.createElement('button');
        upBtn.textContent = '↑';
        upBtn.title = 'Mover para cima';
        upBtn.className = 'move-row-up';
        upBtn.style.padding = '2px 7px';
        upBtn.style.marginRight = '2px';
        upBtn.addEventListener('click', (e) => {
            e.stopPropagation();
            moveUpCallback(row);
        });
        // Botão para baixo
        const downBtn = document.createElement('button');
        downBtn.textContent = '↓';
        downBtn.title = 'Mover para baixo';
        downBtn.className = 'move-row-down';
        downBtn.style.padding = '2px 7px';
        downBtn.addEventListener('click', (e) => {
            e.stopPropagation();
            moveDownCallback(row);
        });
        td.appendChild(upBtn);
        td.appendChild(downBtn);
        return td;
    }

    // --- Função para mover linha na tabela de receitas ---
    function moveRowReceitas(row, direction) {
        const tbody = dataTable.querySelector('tbody');
        const rows = Array.from(tbody.children);
        const idx = rows.indexOf(row);
        if (direction === -1 && idx > 0) {
            tbody.insertBefore(row, rows[idx - 1]);
        } else if (direction === 1 && idx < rows.length - 1) {
            tbody.insertBefore(rows[idx + 1], row);
        }
        saveTableData();
    }

    // --- Função para mover linha na tabela de despesas ---
    function moveRowDespesas(row, direction) {
        const tbody = dataTableExpenses.querySelector('tbody');
        const rows = Array.from(tbody.children);
        const idx = rows.indexOf(row);
        if (direction === -1 && idx > 0) {
            tbody.insertBefore(row, rows[idx - 1]);
        } else if (direction === 1 && idx < rows.length - 1) {
            tbody.insertBefore(rows[idx + 1], row);
        }
        saveTableDataExpenses();
    }
});