# DocTemplate
A simple tool for working with document templates
Is a wrapper over phpWord (https://phpword.readthedocs.io)

## Functions
1. values - Set values from a two-dimensional array in a template document
2. rows - Clones a table row and populates it's values from a two-dimensional array in a template document
3. blocks - Clone block and replace values inside the block to clone
4. table - Insert table in document
5. save - Local save or download document

## General usage example
```php
$rep = new DocTemplate('Sample_40.docx');

$rep->values(['title' => 'Billy', 'inline' => 'Inline Bill!!!']);

$arr = [['cell A1', 'cell A2', 'cell A3'],['cell B1', 'cell B2', 'cell B3']];
$options = [
  'head' => ['column 1','column 2','column 3'],
  'width' => [2000, 4000, 4000],
  'fontSize' => 14,
  'columns' =>[
      ['color' => 'FF0000', 'bold' => 'true', 'alignment' => 'center'],
      ['italic' => true]
    ]
];

$rep->table('table', $arr, $options);

$rep->rows('userId', [['userId' =>1, 'field' => 'Billy 1', 'bield' => True], ['userId' =>10, 'field' => 'Bill!!!', 'bield' => 'Maugli']]);

$rep->blocks('list', [['blink' => 'More more'], ['blink' => 'Gore Bore'], ['blink' => 'Mole Kole'], ]);

$rep->save();
```
