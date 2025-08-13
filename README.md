# Data Transfer Utility (Excel to PostgreSQL)

## Features
- Transfer data from Excel to PostgreSQL with automatic relationship handling
- Supports multiple sheets and relational tables with foreign key lookups
- Uses mapping and relations sheets for flexible configuration
- Automatically handles parent-child table insertion order using DFS algorithm
- Built-in data transformation library (string, int, float, boolean, date, email)
- Lookup cache for performance optimization
- Generated key tracking for cross-table references
- Direct database connection without CLI arguments

## Structure
```
data_transfer_utility_v2/
├── index.js          # Main application with all logic
├── config.js         # Database configuration
├── mapping.xlsx      # Mapping and relations configuration
├── data.xlsx         # Source data file
├── package.json      # Dependencies and project info
└── README.md
```

## Dependencies
- `exceljs`: Excel file processing
- `pg`: PostgreSQL client
- `uuid`: UUID generation
- `xlsx`: Additional Excel support

## Usage
```bash
npm install
node index.js
```

## Configuration
Update database connection in `config.js` or directly in `index.js`:
```javascript
const client = new Client({
    user: 'your_user',
    host: 'localhost',
    database: 'your_database',
    password: 'your_password',
    port: 5432
});
```

## Sheets Explanation
### 1. Mapping Sheet (`mapping`)
| Sheet | SourceColumn | DestinationColumn | Table | DataType | Mandatory |
|-------|---------------|-------------------|--------|----------|-----------|
- **Sheet**: Source Excel sheet name
- **SourceColumn**: Column name in Excel
- **DestinationColumn**: Target database field
- **Table**: Target database table
- **DataType**: Data transformation type (string, int, float, boolean, date, email)
- **Mandatory**: "yes" for required fields

### 2. Relations Sheet (`relations`)
| ParentTable | ChildTable | ForeignKeyColumn | ReferenceColumn | LookupField |
|-------------|-------------|------------------|-----------------|-------------|
- **ParentTable**: Parent table name
- **ChildTable**: Child table name
- **ForeignKeyColumn**: Foreign key field in child table
- **ReferenceColumn**: Primary key field in parent table
- **LookupField**: Field to match values (defaults to 'name')

## Data Transformation Types
- `string`: Trim whitespace
- `int`: Convert to integer
- `float`: Convert to float
- `boolean`: Convert to boolean
- `date`: Convert to ISO date string
- `email`: Convert to lowercase email
- `trim`: Trim whitespace
- `toLowerCase`: Convert to lowercase

## Key Features
- **Automatic Insertion Order**: Uses DFS algorithm to determine correct table insertion sequence
- **Lookup Resolution**: Automatically resolves foreign key relationships using database lookups
- **Generated Key Tracking**: Tracks auto-generated IDs for cross-table references
- **Performance Optimization**: Implements lookup caching to minimize database queries
- **Flexible Mapping**: Supports complex data transformations and relationship mappings