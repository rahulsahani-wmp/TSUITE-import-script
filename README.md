# Data Transfer Utility (Excel to PostgreSQL)

## Features
- Transfer data from Excel to PostgreSQL
- Supports multiple sheets and relational tables
- Uses a mapping sheet for flexible configuration
- Automatically handles parent-child table insertion order
- CLI arguments for dynamic usage
- Validation summary report in Excel
- Logs errors in a separate error file

## Structure
```
data_transfer_utility/
├── index.js
├── config.js
├── utils/
│   ├── logger.js
│   ├── db.js
│   └── excel.js
├── mapping.xlsx
├── data.xlsx
└── README.md
```

## Usage
```bash
node index.js --mapping ./mapping.xlsx --data ./data.xlsx
```

## Sheets Explanation
### 1. Mapping Sheet (`mapping`)
| Sheet | SourceColumn | DestinationColumn | Table | DataType | Mandatory |
|-------|---------------|-------------------|--------|----------|-----------|

### 2. Relations Sheet (`relations`)
| ParentTable | ChildTable | ForeignKeyColumn | ReferenceColumn |
|-------------|-------------|------------------|------------------|

## Outputs
- `error_log.txt` for failed records
- `validation_summary_<timestamp>.xlsx` for validation result