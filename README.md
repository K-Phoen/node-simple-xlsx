XLSX
====

Simple XLSX writer.

# Installation

```
npm install node-simple-xlsx
```

# Usage

```javascript
var XlsxWriter = require('node-simple-xlsx'),
    writer = new XlsxWriter();

writer.addRow({
    'Name': 'Bob',
    'Location': 'Sweden'
});
writer.addRow({
    'Name': 'Alice',
    'Location': 'France'
});
writer.addRow({
    'Name': 'Bob',
    'Location': 'France'
});
writer.addRow({
    'Name': 'Bob',
    'Location': 'France'
});

writer.pack('test.xlsx', function (err) {
    if (err) {
        console.log('Error: ', err);
    } else {
        console.log('Done.');
    }
});
```

# License

This library is released under the MIT license. See the bundled LICENSE file
for details.
