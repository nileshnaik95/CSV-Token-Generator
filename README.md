
### Using Node.js

1. **Install Node.js and npm** if you haven't already.
2. **Install the required packages**:
    ```bash
    npm install csv-parser csv-writer jsonwebtoken
    ```

3. **Create a JavaScript file**:
    ```javascript
    const fs = require('fs');
    const csv = require('csv-parser');
    const createCsvWriter = require('csv-writer').createObjectCsvWriter;
    const jwt = require('jsonwebtoken');

    const secretKey = 'your_secret_key_here';

    const inputCsv = 'input.csv';
    const outputCsv = 'output.csv';

    const records = [];

    fs.createReadStream(inputCsv)
      .pipe(csv())
      .on('data', (row) => {
        if (row.email) {
          row.token = jwt.sign({ email: row.email }, secretKey, { expiresIn: '1d' });
        }
        records.push(row);
      })
      .on('end', () => {
        const csvWriter = createCsvWriter({
          path: outputCsv,
          header: Object.keys(records[0]).map(key => ({ id: key, title: key }))
        });

        csvWriter.writeRecords(records)
          .then(() => {
            console.log('Tokens have been added to the CSV.');
          });
      });
    ```

4. **Run the script**:
    Save the script to a file (e.g., `generate_tokens.js`) and run it:
    ```bash
    node generate_tokens.js
    ```

### Using Excel with VBA

1. **Open your CSV file in Excel**.
2. **Open the VBA editor** (Alt + F11).
3. **Insert a new module** (Right-click on any of the items in the VBA Project panel > Insert > Module).
4. **Add the following VBA code**:
    ```vba
    Sub GenerateTokens()
        Dim ws As Worksheet
        Set ws = ThisWorkbook.Sheets(1)
        
        Dim emailRange As Range
        Set emailRange = ws.Range("A2:A" & ws.Cells(ws.Rows.Count, 1).End(xlUp).Row)
        
        Dim emailCell As Range
        For Each emailCell In emailRange
            If emailCell.Value <> "" Then
                emailCell.Offset(0, 1).Value = GenerateJWT(emailCell.Value)
            End If
        Next emailCell
        
        MsgBox "Tokens have been added to the CSV."
    End Sub

    Function GenerateJWT(email As String) As String
        ' This function simulates JWT generation
        ' Replace this with actual JWT generation logic or use an external library
        GenerateJWT = "jwt_token_for_" & email
    End Function
    ```

5. **Run the macro**:
    Close the VBA editor and run the `GenerateTokens` macro (Alt + F8, select `GenerateTokens`, and click Run).
