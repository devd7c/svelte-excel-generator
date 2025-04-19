/**
 * This file will automatically be loaded by webpack and run in the "renderer" context.
 * To learn more about the differences between the "main" and the "renderer" context in
 * Electron, visit:
 * https://electronjs.org/docs/tutorial/process-model
 */

import * as XLSX from 'xlsx';
import { saveAs } from 'file-saver';
import './index.css';

let providers = [];
let accounts = [];
let branches = [];

let providersMap = new Map();
let accountsMap = new Map();
let branchesMap = new Map();
let tableRows = [];

let debitAccounts = [];
let creditAccounts = [];

let currentRegistros = [];

function appendToGrid(registros) {
    tableRows = [...tableRows, ...registros];
    renderTable();
}

document.getElementById('excelInput').addEventListener('change', (e) => {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (e) => {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });

        providers = XLSX.utils.sheet_to_json(workbook.Sheets['Providers']);
        accounts = XLSX.utils.sheet_to_json(workbook.Sheets['Accounts']);
        branches = XLSX.utils.sheet_to_json(workbook.Sheets['Branches']);

        providersMap = new Map(providers.map(p => [p.id_provider, p.name_provider]));
        accountsMap = new Map(accounts.map(a => [a.id_account.toString(), a.name_account]));
        branchesMap = new Map(branches.map(b => [b.id_branch, b.name_branch]));

        populateSelect('provider', [...providersMap.values()]);
        populateSelect('branch', [...branchesMap.values()]);
    };
    reader.readAsArrayBuffer(file);
});

function populateSelect(id, values) {
    const el = document.getElementById(id);
    el.innerHTML = `<option disabled selected>${el.id}</option>`;
    values.forEach(v => {
        const opt = document.createElement('option');
        opt.textContent = v;
        el.appendChild(opt);
    });
}

document.getElementById('branch').addEventListener('change', () => {
    const branchName = document.getElementById('branch').value;
    const branchId = [...branchesMap.entries()].find(([, v]) => v.trim() === branchName.trim())?.[0] || '';

    const filtered = accounts.filter(acc =>
        (acc.id_branches || '').toString().includes(branchId)
    );

    debitAccounts = filtered.filter(a => a.type_account?.trim().toLowerCase() === 'debit');
    creditAccounts = filtered.filter(a => a.type_account?.trim().toLowerCase() === 'credit');

    populateSelect('debit1', debitAccounts.map(a => a.name_account));
    populateSelect('debit2', debitAccounts.map(a => a.name_account));
    populateSelect('credit3', creditAccounts.map(a => a.name_account));
});

document.getElementById('addRow').addEventListener('click', () => {
    const provider = document.getElementById('provider').value;
    const branch = document.getElementById('branch').value;
    const billNumber = document.getElementById('billNumber').value;
    const expirationDate = document.getElementById('expirationDate').value;
    const billAmount = parseFloat(document.getElementById('billAmount').value);
    const debit1 = document.getElementById('debit1').value;
    const debit2 = document.getElementById('debit2').value;

    const detalle = `F/${billNumber} ${provider}`;
    const vencimiento = expirationDate;

    currentRegistros = [];

    [debit1, debit2].forEach(account => {
        const accId = [...accountsMap.entries()].find(([, v]) => v.trim() === account.trim())?.[0] || '';
        const acc = accounts.find(a => a.name_account.trim() === account.trim());
        if (!acc) return;
        const agencia = ['112070101', '10000000006'].includes(accId) ? '1N' : (
            [...branchesMap.entries()].find(([, v]) => v.trim() === branch.trim())?.[0] || ''
        );
        const valor = billAmount * parseFloat(acc?.percentage_account || 1);

        currentRegistros.push({
            AGENCIA: agencia,
            CUENTA: accId,
            DESCRIPCION: acc?.name_account || '',
            DOCUMENTO: accId === '112070101' ? '' : acc.code_account || '',
            VENCIMIENTO: accId === '10000000006' ? vencimiento : '',
            DETALLE: accId === '112070101' ? '' : detalle,
            DEBEBOL: formatAmount(valor),
            HABERBOL: 0,
            DEBEDOL: '',
            HABERDOL: '',
            REFERENCIA: ['112070101', '10000000006'].includes(accId) ? '' : acc.code_account || ''
        });
    });

    const providerId = [...providersMap.entries()].find(([, v]) => v.trim() === provider.trim())?.[0] || '';
    const acc = accounts.find(a => a.id_account.toString() === '10000000006');
    const valor = billAmount * parseFloat(acc?.percentage_account || 1);

    currentRegistros.push({
        AGENCIA: '1N',
        CUENTA: providerId,
        DESCRIPCION: provider,
        DOCUMENTO: `F/${billNumber}`,
        VENCIMIENTO: vencimiento,
        DETALLE: detalle,
        DEBEBOL: 0,
        HABERBOL: formatAmount(valor),
        DEBEDOL: '',
        HABERDOL: '',
        REFERENCIA: ''
    });

    appendToGrid(currentRegistros);
});

function renderTable() {
    const table = document.createElement('table');
    table.className = 'table-auto w-full text-xs';

    table.innerHTML = `
    <thead><tr>
      <th>AGENCIA</th><th>CUENTA</th><th>DESCRIPCIÓN</th><th>DOCUMENTO</th><th>VENCIMIENTO</th>
      <th>DETALLE</th><th>DEBEBOL</th><th>HABERBOL</th><th>DEBEDOL</th><th>HABERDOL</th><th>REFERENCIA</th>
    </tr></thead>
    <tbody>
      ${tableRows.map(row => `
        <tr>
          <td>${row.AGENCIA}</td>
          <td>${row.CUENTA}</td>
          <td>${row.DESCRIPCION}</td>
          <td>${row.DOCUMENTO}</td>
          <td>${row.VENCIMIENTO}</td>
          <td>${row.DETALLE}</td>
          <td>${row.DEBEBOL}</td>
          <td>${row.HABERBOL}</td>
          <td>${row.DEBEDOL}</td>
          <td>${row.HABERDOL}</td>
          <td>${row.REFERENCIA}</td>
        </tr>
      `).join('')}
    </tbody>
  `;

    const container = document.getElementById('tableContainer');
    container.innerHTML = '';
    container.appendChild(table);
}

function formatAmount(val) {
    const rounded = Number(val.toFixed(2));
    return rounded % 1 === 0 ? rounded.toFixed(0) : rounded.toFixed(2);
}

document.getElementById('exportExcel').addEventListener('click', () => {
    const ws = XLSX.utils.json_to_sheet(tableRows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Grid');
    const blob = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    saveAs(new Blob([blob]), 'grid_export.xlsx');
});