// This file contains the main JavaScript logic for the application, including functions for managing a table-like interface, handling user interactions, and managing data storage.

document.addEventListener('DOMContentLoaded', () => {
    const addRowButton = document.getElementById('addRow');
    const addColumnButton = document.getElementById('addColumn');
    const dataTable = document.getElementById('dataTable');

    const loadTableData = () => {
        const savedData = JSON.parse(localStorage.getItem('tableData')) || { headers: [], rows: [] };
        const thead = dataTable.querySelector('thead tr');
        const tbody = dataTable.querySelector('tbody');

        thead.innerHTML = '';
        savedData.headers.forEach(header => {
            const th = document.createElement('th');
            th.textContent = header.name;
            th.dataset.type = header.type;
            thead.appendChild(th);
        });

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
        const colType = document.getElementById('colType').value;
        const thead = dataTable.querySelector('thead');
        const tbody = dataTable.querySelector('tbody');

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

        saveTableData();
    });

    let selectedRows = new Set();
    let selectedColumns = new Set();

    dataTable.addEventListener('click', (event) => {
        const target = event.target;

        if (target.tagName === 'TD') {
            const row = target.parentElement;
            const columnIndex = Array.from(row.children).indexOf(target);

            if (event.ctrlKey) {
                if (selectedRows.has(row)) {
                    selectedRows.delete(row);
                    row.classList.remove('selected');
                } else {
                    selectedRows.add(row);
                    row.classList.add('selected');
                }
            } else {
                selectedRows.forEach(r => r.classList.remove('selected'));
                selectedColumns.forEach(index => {
                    dataTable.querySelectorAll('tr').forEach(row => {
                        const cell = row.children[index];
                        if (cell) cell.classList.remove('selected');
                    });
                });
                selectedRows.clear();
                selectedColumns.clear();

                selectedRows.add(row);
                row.classList.add('selected');
            }
        } else if (target.tagName === 'TH') {
            const columnIndex = Array.from(target.parentElement.children).indexOf(target);

            if (event.ctrlKey) {
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
                selectedRows.forEach(r => r.classList.remove('selected'));
                selectedColumns.forEach(index => {
                    dataTable.querySelectorAll('tr').forEach(row => {
                        const cell = row.children[index];
                        if (cell) cell.classList.remove('selected');
                    });
                });
                selectedRows.clear();
                selectedColumns.clear();

                selectedColumns.add(columnIndex);
                dataTable.querySelectorAll('tr').forEach(row => {
                    const cell = row.children[columnIndex];
                    if (cell) cell.classList.add('selected');
                });
            }
        } else {
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
    deleteSelectionButton.style.display = 'none';
    document.querySelector('.container').appendChild(deleteSelectionButton);

    deleteSelectionButton.addEventListener('click', () => {
        selectedRows.forEach(row => row.remove());
        selectedRows.clear();

        Array.from(selectedColumns).sort((a, b) => b - a).forEach(index => {
            const thead = dataTable.querySelector('thead tr');
            const tbody = dataTable.querySelector('tbody');

            if (thead.children[index]) thead.children[index].remove();

            tbody.querySelectorAll('tr').forEach(row => {
                if (row.children[index]) row.children[index].remove();
            });
        });
        selectedColumns.clear();

        deleteSelectionButton.style.display = 'none';
        saveTableData();
    });

    dataTable.addEventListener('dblclick', (event) => {
        const target = event.target;

        if (target.tagName === 'TH') {
            const currentText = target.textContent;
            const input = document.createElement('input');
            input.type = 'text';
            input.value = currentText;
            input.style.width = '100%';

            target.textContent = '';
            target.appendChild(input);

            input.addEventListener('blur', () => {
                target.textContent = input.value || currentText;
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

    dataTable.addEventListener('dblclick', (event) => {
        const target = event.target;

        if (target.tagName === 'TD') {
            const currentText = target.textContent;
            const input = document.createElement('input');
            input.type = 'text';
            input.value = currentText;
            input.style.width = '100%';

            target.textContent = '';
            target.appendChild(input);

            input.addEventListener('blur', () => {
                target.textContent = input.value || currentText;
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

    document.addEventListener('keydown', (event) => {
        if (event.key === 'Escape') {
            dataTable.querySelectorAll('tr').forEach(row => row.classList.remove('selected'));
            dataTable.querySelectorAll('th, td').forEach(cell => cell.classList.remove('selected'));
            selectedRows.clear();
            selectedColumns.clear();
            deleteSelectionButton.style.display = 'none';
        }
    });

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

    loadTableData();
});