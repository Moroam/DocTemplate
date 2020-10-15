<?php
error_reporting(E_ALL);
if(ini_set('display_errors', 1)===false)
  echo "ERROR INI SET";

require '/vendor/autoload.php';

include_once 'doctemplate.class.php';

$rep = new DocTemplate('Sample.docx');

$rep->values(['title' => 'Billy', 'inline' => 'Inline Bill!!!']);

$arr = [['cell A1', 'cell A2', 'cell A3'],['cell B1', 'cell B2', 'cell B3']];
$options = [
  'head' => ['column 1','column 2','column 3'],
  'width' => [1, 2, 2],
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
