document.addEventListener('DOMContentLoaded', () => {
    const addRowButton = document.getElementById('addRow');
    const addColumnButton = document.getElementById('addColumn');
    const dataTable = document.getElementById('dataTable');

    // Load table data from localStorage
    const loadTableData = () => {
        const savedData = JSON.parse(localStorage.getItem('tableData')) || { headers: [], rows: [] };
        const thead = dataTable.querySelector('thead tr');
        const tbody = dataTable.querySelector('tbody');

        // Populate headers
        thead.innerHTML = '';
        savedData.headers.forEach(header => {
            const th = document.createElement('th');
            th.textContent = header.name;
            th.dataset.type = header.type;
            thead.appendChild(th);
        });

        // Populate rows
        tbody.innerHTML = '';
        savedData.rows.forEach(rowData => {
            const tr = document.createElement('tr');
            rowData.forEach(cellData => {
                const td = document.createElement('td');
                td.textContent = cellData;
                tr.appendChild(td);
            });
            tbody.appendChild(tr);
        });

        updateSumResult();
    };

    // Save table data to localStorage
    const saveTableData = () => {
        const headers = Array.from(dataTable.querySelector('thead tr').children).map(th => ({
            name: th.textContent,
            type: th.dataset.type || 'text'
        }));
        const rows = Array.from(dataTable.querySelector('tbody').children).map(tr =>
            Array.from(tr.children).map(td => td.textContent)
        );
        localStorage.setItem('tableData', JSON.stringify({ headers, rows }));
        updateSumResult();
    };

    addRowButton.addEventListener('click', () => {
        const tbody = dataTable.querySelector('tbody');
        const newRow = document.createElement('tr');
        const columnCount = dataTable.querySelector('thead tr').children.length;

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
        newHeader.textContent = colName || `Column ${headerRow.children.length + 1}`;
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
            const currentText = target.textContent;
            const input = document.createElement('input');
            input.type = 'text';
            input.value = currentText;
            input.style.width = '100%';

            // Replace cell text with input field
            target.textContent = '';
            target.appendChild(input);

            // Save changes on blur or Enter key
            input.addEventListener('blur', () => {
                target.textContent = input.value || currentText; // Revert if empty
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
                    const value = parseFloat(cell.textContent) || 0;
                    sum += value;
                }
            });
        });

        document.getElementById('sumResult').textContent = sum.toFixed(2);
    };

    // Load table data on page load
    loadTableData();

    const addRowButtonExpenses = document.getElementById('addRowExpenses');
    const addColumnButtonExpenses = document.getElementById('addColumnExpenses');
    const dataTableExpenses = document.getElementById('dataTableExpenses');

    const loadTableDataExpenses = () => {
        const savedData = JSON.parse(localStorage.getItem('tableDataExpenses')) || { headers: [], rows: [] };
        const thead = dataTableExpenses.querySelector('thead tr');
        const tbody = dataTableExpenses.querySelector('tbody');

        // Populate headers
        thead.innerHTML = '';
        savedData.headers.forEach(header => {
            const th = document.createElement('th');
            th.textContent = header.name;
            th.dataset.type = header.type;
            thead.appendChild(th);
        });

        // Populate rows
        tbody.innerHTML = '';
        savedData.rows.forEach(rowData => {
            const tr = document.createElement('tr');
            rowData.forEach(cellData => {
                const td = document.createElement('td');
                td.textContent = cellData;
                tr.appendChild(td);
            });
            tbody.appendChild(tr);
        });

        updateSumResultExpenses();
    };

    const saveTableDataExpenses = () => {
        const headers = Array.from(dataTableExpenses.querySelector('thead tr').children).map(th => ({
            name: th.textContent,
            type: th.dataset.type || 'text'
        }));
        const rows = Array.from(dataTableExpenses.querySelector('tbody').children).map(tr =>
            Array.from(tr.children).map(td => td.textContent)
        );
        localStorage.setItem('tableDataExpenses', JSON.stringify({ headers, rows }));
        updateSumResultExpenses();
    };

    addRowButtonExpenses.addEventListener('click', () => {
        const tbody = dataTableExpenses.querySelector('tbody');
        const newRow = document.createElement('tr');
        const columnCount = dataTableExpenses.querySelector('thead tr').children.length;

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
        newHeader.textContent = colName || `Column ${headerRow.children.length + 1}`;
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
                    const value = parseFloat(cell.textContent) || 0;
                    sum += value;
                }
            });
        });

        document.getElementById('sumResultExpenses').textContent = sum.toFixed(2);
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
            const currentText = target.textContent;
            const input = document.createElement('input');
            input.type = 'text';
            input.value = currentText;
            input.style.width = '100%';

            // Replace cell text with input field
            target.textContent = '';
            target.appendChild(input);

            // Save changes on blur or Enter key
            input.addEventListener('blur', () => {
                target.textContent = input.value || currentText; // Revert if empty
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
        const receitas = parseFloat(document.getElementById('sumResult').textContent) || 0;
        const despesas = parseFloat(document.getElementById('sumResultExpenses').textContent) || 0;
        const resultadoFinal = receitas - despesas;
        document.getElementById('finalResult').textContent = resultadoFinal.toFixed(2);
    });

    // Limpar Tudo
    document.getElementById('clearAll').addEventListener('click', () => {
        document.getElementById('finalResult').textContent = '0';
    });

    // Exportar Dados (agora em formato Excel)
    document.getElementById('exportData').addEventListener('click', () => {
        // Receitas
        const receitasData = JSON.parse(localStorage.getItem('tableData')) || { headers: [], rows: [] };
        const receitasSheet = [
            receitasData.headers.map(h => h.name),
            ...receitasData.rows
        ];

        // Despesas
        const despesasData = JSON.parse(localStorage.getItem('tableDataExpenses')) || { headers: [], rows: [] };
        const despesasSheet = [
            despesasData.headers.map(h => h.name),
            ...despesasData.rows
        ];

        // Cria workbook
        const wb = XLSX.utils.book_new();
        const wsReceitas = XLSX.utils.aoa_to_sheet(receitasSheet);
        XLSX.utils.book_append_sheet(wb, wsReceitas, "Receitas");
        const wsDespesas = XLSX.utils.aoa_to_sheet(despesasSheet);
        XLSX.utils.book_append_sheet(wb, wsDespesas, "Despesas");

        // Salva arquivo
        XLSX.writeFile(wb, "dados-financeiros.xlsx");
    });

    // Importar Dados (agora aceita .xlsx exportado)
    document.getElementById('importData').addEventListener('click', () => {
        document.getElementById('fileInput').accept = '.xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
        document.getElementById('fileInput').click();
    });

    document.getElementById('fileInput').addEventListener('change', (event) => {
        const file = event.target.files[0];
        if (!file) return;

        // Verifica extensão
        if (file.name.endsWith('.xlsx')) {
            const reader = new FileReader();
            reader.onload = (e) => {
                try {
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, { type: 'array' });

                    // Receitas
                    const wsReceitas = workbook.Sheets['Receitas'];
                    const receitasSheet = wsReceitas ? XLSX.utils.sheet_to_json(wsReceitas, { header: 1 }) : [];
                    let receitas = { headers: [], rows: [] };
                    if (receitasSheet.length > 0) {
                        receitas.headers = receitasSheet[0].map(name => ({ name, type: name.toLowerCase().includes('valor') ? 'value' : 'text' }));
                        receitas.rows = receitasSheet.slice(1);
                    }

                    // Despesas
                    const wsDespesas = workbook.Sheets['Despesas'];
                    const despesasSheet = wsDespesas ? XLSX.utils.sheet_to_json(wsDespesas, { header: 1 }) : [];
                    let despesas = { headers: [], rows: [] };
                    if (despesasSheet.length > 0) {
                        despesas.headers = despesasSheet[0].map(name => ({ name, type: name.toLowerCase().includes('valor') ? 'value' : 'text' }));
                        despesas.rows = despesasSheet.slice(1);
                    }

                    localStorage.setItem('tableData', JSON.stringify(receitas));
                    localStorage.setItem('tableDataExpenses', JSON.stringify(despesas));
                    // Recarrega tabelas
                    loadTableData();
                    loadTableDataExpenses();
                    document.getElementById('finalResult').textContent = '0';
                } catch {
                    alert('Erro ao importar arquivo Excel.');
                }
            };
            reader.readAsArrayBuffer(file);
        } else {
            alert('Por favor, selecione um arquivo .xlsx exportado pelo sistema.');
        }
        // Limpa input para permitir novo upload igual
        event.target.value = '';
    });

    // Importar Dados
    document.getElementById('importData').addEventListener('click', () => {
        document.getElementById('fileInput').click();
    });

    document.getElementById('fileInput').addEventListener('change', (event) => {
        const file = event.target.files[0];
        if (!file) return;
        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const data = JSON.parse(e.target.result);
                if (data.receitas && data.despesas) {
                    localStorage.setItem('tableData', JSON.stringify(data.receitas));
                    localStorage.setItem('tableDataExpenses', JSON.stringify(data.despesas));
                    // Recarrega tabelas
                    loadTableData();
                    loadTableDataExpenses();
                    document.getElementById('finalResult').textContent = '0';
                } else {
                    alert('Arquivo inválido.');
                }
            } catch {
                alert('Erro ao importar dados.');
            }
        };
        reader.readAsText(file);
        // Limpa input para permitir novo upload igual
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

    // ...existing code...
});