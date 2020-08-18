<h2>Definitions:</h2>
<p><b>Parent Column:</b> The first column of a given range
<p><b>Secondary Column:</b> Any columns of a given range, excluding the parent column
<h2>Module Description</h2>
<p>This module will parse through each row of a given range, returning the first non-blank, non-zero value it finds. The corresponding cell in the parent column is changed to this returned value. Once all rows have been parsed, all secondary columns will be deleted, leaving only the parent column.
