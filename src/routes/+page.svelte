<script lang="ts">
    import { onMount } from 'svelte';
    import * as XLSX from 'xlsx';
    import pkg from 'file-saver';
    const { saveAs } = pkg

    let providers: any[] = [];
    let accounts: any[] = [];
    let branches: any[] = [];

    let providersMap = new Map();
    let accountsMap = new Map();
    let branchesMap = new Map();

    let provider = '';
    let branch = '';
    let billNumber = '';
    let expirationDate = '';
    let billAmount = '';
    let debit1 = '';
    let debit2 = '';
    let credit3 = '';

    let tableRows: any[] = [];

    let debitAccounts: string[] = [];
    let creditAccounts: string[] = [];

    function parseExcel(file: File) {
        const reader = new FileReader();
        reader.onload = (e) => {
            const data = new Uint8Array(e.target?.result as ArrayBuffer);
            const workbook = XLSX.read(data, { type: 'array' });
            providers = XLSX.utils.sheet_to_json(workbook.Sheets['Providers']);
            accounts = XLSX.utils.sheet_to_json(workbook.Sheets['Accounts']);
            branches = XLSX.utils.sheet_to_json(workbook.Sheets['Branches']);

            providersMap = new Map(providers.map(p => [p.id_provider, p.name_provider]));
            accountsMap = new Map(accounts.map(a => [a.id_account.toString(), a.name_account]));
            branchesMap = new Map(branches.map(b => [b.id_branch, b.name_branch]));
        };
        reader.readAsArrayBuffer(file);
    }

    function onBranchChange() {
        const branchId = [...branchesMap.entries()].find(([, v]) => v === branch)?.[0] || '';

        const filtered = accounts.filter(acc =>
            (acc.id_branches || '').toString().includes(branchId)
        );

        debitAccounts = filtered.filter(a => a.type_account?.trim().toLowerCase() === 'debit').map(a => a.name_account);
        creditAccounts = filtered.filter(a => a.type_account?.trim().toLowerCase() === 'credit').map(a => a.name_account);
    }

    function formatAmount(val: number) {
        const rounded = Number(val.toFixed(2));
        return rounded % 1 === 0 ? rounded.toFixed(0) : rounded.toFixed(2);
    }

    function addToGrid() {
        const getAccountRow = (accountName: string) => {
            const accId = [...accountsMap.entries()].find(([, v]) => v === accountName)?.[0] || '';
            const acc = accounts.find(a => a.id_account.toString() === accId);
            return { accId, acc };
        };

        const vencimiento = expirationDate;
        const detalle = `F/${billNumber} ${provider}`;
        const bill = parseFloat(billAmount);

        const registros: any[] = [];

        [debit1, debit2].forEach(account => {
            const { accId, acc } = getAccountRow(account);
            const agencia = ['112070101', '10000000006'].includes(accId) ? '1N' : ([...branchesMap.entries()].find(([, v]) => v === branch)?.[0] || '');
            const valor = bill * parseFloat(acc.percentage_account);
            registros.push({
                AGENCIA: agencia,
                CUENTA: accId,
                DESCRIPCION: acc.name_account,
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

        // Credito
        const providerId = [...providersMap.entries()].find(([, v]) => v === provider)?.[0] || '';
        const acc = accounts.find(a => a.id_account.toString() === '10000000006');
        const valor = bill * parseFloat(acc.percentage_account);
        registros.push({
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

        tableRows = [...tableRows, ...registros];
    }

    function exportToExcel() {
        const ws = XLSX.utils.json_to_sheet(tableRows);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, 'Grid');
        const blob = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
        saveAs(new Blob([blob]), 'grid_export.xlsx');
    }
</script>

<input type="file" accept=".xlsx" on:change={(e) => parseExcel(e.target.files[0])} class="mb-4" />

<div class="grid grid-cols-1 gap-4">
    <div class="grid grid-cols-2 gap-4">
        <select bind:value={provider} class="border p-2">
            <option disabled selected>Proveedor</option>
            {#each [...providersMap.values()] as name}
                <option>{name}</option>
            {/each}
        </select>

        <select bind:value={branch} class="border p-2" on:change={onBranchChange}>
            <option disabled selected>Agencia</option>
            {#each [...branchesMap.values()] as name}
                <option>{name}</option>
            {/each}
        </select>
    </div>

    <div class="grid grid-cols-3 gap-4">
        <input type="text" placeholder="Nro Factura" bind:value={billNumber} class="border p-2" />
        <input type="date" bind:value={expirationDate} class="border p-2" />
        <input type="number" step="0.01" placeholder="Monto" bind:value={billAmount} class="border p-2" />
    </div>

    <div class="grid grid-cols-3 gap-4">
        <select bind:value={debit1} class="border p-2">
            <option disabled selected>Cuenta Débito 1</option>
            {#each debitAccounts as acc}
                <option>{acc}</option>
            {/each}
        </select>
        <select bind:value={debit2} class="border p-2">
            <option disabled selected>Cuenta Débito 2</option>
            {#each debitAccounts as acc}
                <option>{acc}</option>
            {/each}
        </select>
        <select bind:value={credit3} class="border p-2">
            <option disabled selected>Cuenta Crédito 3</option>
            {#each creditAccounts as acc}
                <option>{acc}</option>
            {/each}
        </select>
    </div>

    <div class="flex gap-4">
        <button class="bg-green-500 text-white px-4 py-2 rounded" on:click={addToGrid}>Agregar a Grid</button>
        <button class="bg-blue-500 text-white px-4 py-2 rounded" on:click={exportToExcel}>Exportar a Excel</button>
    </div>

    {#if tableRows.length}
        <div class="overflow-x-auto mt-6 border p-4 rounded shadow">
            <table class="table-auto w-full text-xs">
                <thead>
                <tr>
                    <th>AGENCIA</th><th>CUENTA</th><th>DESCRIPCIÓN</th><th>DOCUMENTO</th><th>VENCIMIENTO</th>
                    <th>DETALLE</th><th>DEBEBOL</th><th>HABERBOL</th><th>DEBEDOL</th><th>HABERDOL</th><th>REFERENCIA</th>
                </tr>
                </thead>
                <tbody>
                {#each tableRows as row}
                    <tr>
                        <td>{row.AGENCIA}</td>
                        <td>{row.CUENTA}</td>
                        <td>{row.DESCRIPCION}</td>
                        <td>{row.DOCUMENTO}</td>
                        <td>{row.VENCIMIENTO}</td>
                        <td>{row.DETALLE}</td>
                        <td>{row.DEBEBOL}</td>
                        <td>{row.HABERBOL}</td>
                        <td>{row.DEBEDOL}</td>
                        <td>{row.HABERDOL}</td>
                        <td>{row.REFERENCIA}</td>
                    </tr>
                {/each}
                </tbody>
            </table>
        </div>
    {/if}
</div>

<style>
    select, input {
        width: 100%;
    }
</style>
