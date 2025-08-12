const { log } = require('console');
const ExcelJS = require('exceljs');
const path = require('path');
const { Client } = require('pg'); // example: for Postgres, replace with your DB client
// ---------------- DB Config ----------------
const client = new Client({
    user: 'teamsuite',
    host: 'localhost',
    database: '17march',
    password: '12345',
    port: 5432
});
// ---------------- Transformation Library ----------------
const transformLib = {
    string: val => String(val || '').trim(),
    int: val => parseInt(val, 10) || 0,
    float: val => parseFloat(val) || 0.0,
    boolean: val => val === true || val === 'true' || val === 1 || val === '1',
    date: val => val ? new Date(val).toISOString() : null,
    email: val => String(val || '').toLowerCase(),
    trim: val => String(val || '').trim(),
    toLowerCase: val => String(val || '').toLowerCase()
};
// ---------------- Load Mapping + Relations ----------------
async function loadMappingSheets(mappingFilePath) {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(mappingFilePath);
    const mappingSheet = workbook.getWorksheet('mapping');
    const relationSheet = workbook.getWorksheet('relations');
const mapping = mappingSheet.getSheetValues().slice(1).map(r => ({
    sheetName: r[1],            // Sheet
    sourceColumn: r[2],         // SourceColumn
    targetField: r[3],          // DestinationColumn
    collection: r[4],           // Table
    datatype: r[5],             // DataType
    mandatory: (r[6] || '').toString().toLowerCase() === 'yes'
}));


    const relations = relationSheet
        ? relationSheet.getSheetValues().slice(2).map(r => ({
            parent: r[1],
            child: r[2],
            childFK: r[3],
            referenceColumn: r[4],
            lookupField: r[5] || 'name'
        }))
        : [];



// Add lookup info from relations
mapping.forEach(map => {
    const relation = relations.find(r => 
        r.child === map.collection && r.childFK === map.targetField
    );
    if (relation) {
        map.lookupCollection = relation.parent;
        map.lookupField = relation.lookupField;
        map.returnField = relation.referenceColumn;
    }
});


    return { mapping, relations };
}
// ---------------- Build Table Graph ----------------
function buildTableGraph(relations) {
    const graph = {};
    relations.forEach(({ parent, child }) => {
        if (!graph[parent]) graph[parent] = [];
        if (!graph[child]) graph[child] = [];
        graph[parent].push(child);
    });
    return graph;
}
function dfsOrder(graph) {
    const visited = new Set();
    const order = [];
    function dfs(node) {
        if (visited.has(node)) return;
        visited.add(node);
        (graph[node] || []).forEach(dfs);
        order.push(node);
    }
    Object.keys(graph).forEach(dfs);
    return order.reverse();
}
// ---------------- Lookup Cache ----------------
const lookupCache = {};
async function getLookupMap(collection, field, returnField) {
    const key = `${collection}_${field}_${returnField}`;
    if (lookupCache[key]) return lookupCache[key];
    
    const query = `SELECT ${field} as lookup_key, ${returnField} FROM ${collection}`;
    const res = await client.query(query);
    const map = {};
    res.rows.forEach(r => {
        map[String(r.lookup_key).trim()] = r[returnField];
    });
    lookupCache[key] = map;
    return map;
}
// ---------------- Process Data Excel ----------------
async function parseDataExcel(dataFilePath, mappingConfig, generatedKeys) {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(dataFilePath);
    const results = {};
    
    for (const sheetName of [...new Set(mappingConfig.map(m => m.sheetName))]) {
        const sheet = workbook.getWorksheet(sheetName);
        if (!sheet) continue;
        
        // Get headers, skip first null value from ExcelJS
        const headerRow = sheet.getRow(1).values.slice(1);
        const headers = headerRow.map(h => String(h || '').trim());
        const mappingForSheet = mappingConfig.filter(m => m.sheetName === sheetName);
        
        // Process each row synchronously
        for (let rowIndex = 2; rowIndex <= sheet.rowCount; rowIndex++) {
            const row = sheet.getRow(rowIndex);
            if (!row.hasValues) continue;
            
            // Get row values, skip first null value from ExcelJS
            const rowValues = row.values.slice(1);
            
            let record = {};
            let collectionName = null;
            
            for (const map of mappingForSheet) {
                const colIndex = headers.indexOf(map.sourceColumn);
                if (colIndex > -1 && colIndex < rowValues.length) {
                    let value = rowValues[colIndex];
                    
                    // Transform
                    if (map.transform && transformLib[map.transform]) {
                        value = transformLib[map.transform](value);
                    } else if (transformLib[map.datatype]) {
                        value = transformLib[map.datatype](value);
                    }
                    
                    // Lookup from DB using relations
                    if (map.lookupCollection && map.lookupField && map.returnField && value) {
                        try {
                            const lookupMap = await getLookupMap(map.lookupCollection, map.lookupField, map.returnField);
                            const lookupValue = lookupMap[String(value).trim()];
                            if (lookupValue !== undefined) {
                                value = lookupValue;
                            }
                        } catch (err) {
                            // Lookup failed, keep original value
                        }
                    }
                    
                    // Lookup from generated keys if child FK
                    const originalValue = String(value).trim();
                    if (generatedKeys[map.lookupCollection]?.[originalValue]) {
                        value = generatedKeys[map.lookupCollection][originalValue];
                    }
                    
                    record[map.targetField] = value;
                    collectionName = map.collection;
                }
            }
            
            if (collectionName && Object.keys(record).length > 0) {
                if (!results[collectionName]) results[collectionName] = [];
                results[collectionName].push(record);
            }
        }
    }
    
 console.log('Parsed results:', results);
    return results;
}
// ---------------- Insert Data in Order ----------------
async function insertInOrder(order, results) {
    const generatedKeys = {};
    for (const table of order) {
        if (!results[table] || results[table].length === 0) continue;
        generatedKeys[table] = {};
        for (const row of results[table]) {
            const cols = Object.keys(row);
            const vals = Object.values(row);
            const placeholders = cols.map((_, i) => `$${i + 1}`).join(',');
            const queryText = `INSERT INTO ${table} (${cols.join(',')}) VALUES (${placeholders}) RETURNING id`;

const res = await client.query(queryText, vals);
            const id = res.rows[0].id;
            // Store generated keys for lookups - use multiple possible key fields
            const keyField = cols.find(c => c.endsWith('name')) || cols.find(c => c.includes('name')) || cols[0];
            if (row[keyField]) {
                generatedKeys[table][String(row[keyField]).trim()] = id;
            }
        }
    }
    return generatedKeys;
}
// ---------------- Main ----------------
(async () => {
    try {
        await client.connect();
        const { mapping, relations } = await loadMappingSheets(path.join(__dirname, 'mapping.xlsx'));
        const graph = buildTableGraph(relations);
        const order = dfsOrder(graph);

        // Initial pass: no generated keys yet
        let parsedResults = await parseDataExcel(path.join(__dirname, 'data.xlsx'), mapping, {});
        
        const generatedKeys = await insertInOrder(order, parsedResults);

        console.log('Data transfer completed successfully!');
        await client.end();
    } catch (err) {
        console.error('Error in import', err);
    }
})();