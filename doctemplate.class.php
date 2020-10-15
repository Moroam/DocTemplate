<?php
use PhpOffice\PhpWord\Element\Table;
class DocTemplate {
  protected $templateProcessor = null;

  protected $tabHeadCellStyle = [
    'borderSize' => 0,
    'valign' => 'center',
    'align' => 'center',
  ];
  protected $tabFirstRowStyle = [
    'bgColor' => 'DDDDDD',
  ];

  # these keys can be used when formatting a table
  const KEYS_STYLE = [
    'cell' => ['valign', 'bgColor'],
    'paragraph' => ['alignment'],
    'font' => ['size', 'color', 'bold', 'italic', 'underline']
  ];

  public function __construct(string $template_path){
    if(!is_file($template_path)){
      echo "<h3>ERROR: template $template_path don't exists";
      return;
    }
    $this->templateProcessor = new \PhpOffice\PhpWord\TemplateProcessor($template_path);
  }

  /**
   * Get templateProcessor
   */
  public function __get(string $name){
    if($name === 'templateProcessor'){
      return $this->templateProcessor;
    }

    return null;
  }

  /**
   * Set values from a two-dimensional array in a template document
   *
   * @param array @values => [['varname1' =. 'value 1'], ['varname1' =. 'value 2'], ...]
   */
  public function values(array $values){
    $this->templateProcessor->setValues($values);
  }

  /**
   * Clones a table row and populates it's values from a two-dimensional array in a template document
   */
  public function rows(string $var_name, array $values){
    $this->templateProcessor->cloneRowAndSetValues($var_name, $values);
  }

  /**
   * Clone block
   *
   * @param string $block_name
   * @param array $replacements Array containing replacements for macros found inside the block to clone
   */
  public function blocks(string $block_name, array $replacements){
    $this->templateProcessor->cloneBlock($block_name, 0, true, false, $replacements);
  }

  /**
   * Recalc cells width for table from $options
   */
  protected function recalc_width(array $options, int $cols){
    if(count($options['width'])==0){
      return array_fill(0, $cols, (int)($options['tabWidth'] / $cols));
    } else {
      for($arr=[],$s=0,$i=0;$i<$cols;$i++){
        $s += $arr[] = $options['width'][$i] ?? 0;
      }
      if($s == 0){
        return array_fill(0, $cols, (int)($options['tabWidth'] / $cols));
      }
      return array_map(fn($x)=>$x==0?1:(int)($options['tabWidth']*$x/$s), $arr); // 0 - invalid width for cell is
    }
  }

  /**
   * Calculate style for cells (cellStyle, paragraphStyle, fontStyle) from $options
   */
  protected function calc_style(array $options, int $cols) : array {
    $cellStyle = [];
    $paragraphStyle = [];
    $fontStyle = [];

    $ma = fn($arr, $frmt, $keys_name) => array_merge($arr, array_filter($frmt, fn($k) => in_array($k, self::KEYS_STYLE[$keys_name]), ARRAY_FILTER_USE_KEY) );

    $columns = $options['columns'];

    for($i = 0; $i < $cols; $i++){
      $cellStyle[] = ['borderSize' => 0];
      $paragraphStyle[] = [];
      $fontStyle[] = ['size' => $options['fontSize']];
      if(array_key_exists($i, $columns)){
        $cellStyle[$i]      = $ma($cellStyle[$i],      $columns[$i], 'cell'     );
        $paragraphStyle[$i] = $ma($paragraphStyle[$i], $columns[$i], 'paragraph');
        $fontStyle[$i]      = $ma($fontStyle[$i],      $columns[$i], 'font'     );
      }
    }

    return ['cellStyle' => $cellStyle, 'paragraphStyle' => $paragraphStyle, 'fontStyle' => $fontStyle];
  }

  /**
   * Insert table in document
   *
   * @param string $block_table_name
   * @param array $data
   * @param array $options = [
   *    array  head     => values for first / head row
   *    array  width    => cells width
   *    string caption => table caption
   *    int    fontSize   => default font size for table
   *    int    tabWidth   => default = tabStyle['width'] (100% page width - margin)
   *    array  columns  => format for cells, may be set for individual cell column_number =>
   *             [
   *              'valign', 'bgColor', #### CELL STYLE
   *              'alignment',         #### PARAGRAPH STYLE
   *              'size', 'color', 'bold', 'italic', 'underline' ### FONT STYLE
   *             ]
   *  ]
   * Parameters value
   * 'valign' => Vertical alignment, top, center, both, bottom
   * 'alignment' => paragraph text alignment 'start', 'center', 'end', 'both', 'left', 'right', 'justify'
   * 'bold', 'italic' => true / false
   * 'underline' => 'none','dash','dashLong','dashLongHeavy','dotDash','dotDotDash','dotted','dottedHeavy','single','wavyHeavy','words'
   */
  public function table(string $block_table_name, array $data, array $options = []){
    if(count($data) == 0) return;

    // INIT OPTIONS
    $options['head'    ] = $options['head'    ] ?? [];
    $options['width'   ] = $options['width'   ] ?? [];
    $options['fontSize'] = $options['fontSize'] ?? 10;
    $options['tabWidth'] = $options['tabWidth'] ?? 10200;
    $options['columns' ] = $options['columns' ] ?? [];

    $cols = count($data[0]);
    $cols_head = count($options['head']);
    if($cols_head > 0 && $cols_head <> $cols){
      echo "<h3>Количество столбцов заголовка таблицы не равно количеству столбцов таблицы с данными</h3>";
      return;
    }
    if(count($options['width'])>0 && count($options['width'])<>$cols){
      echo "<h3>Количество столбцов таблицы с шириной столбцов не равно количеству столбцов таблицы с данными</h3>";
      return;
    }

    $options['width'] = $this->recalc_width($options, $options['tabWidth'], $cols);
    $style = $this->calc_style($options, $cols);

    // add table
    if($cols_head == 0){
      $table = new Table(['borderSize' => 0,'cellMargin' => 0,'alignment'  => 'center']);
    } else {
      // add head
      $table = new Table(['borderSize' => 0,'cellMargin' => 0,'alignment'  => 'center'], ['bgColor' => 'DDDDDD']);
      $table->addRow();
      foreach ($options['head'] as $k => $val) {
        $table
          ->addCell($options['width'][$k] ?? 0, $this->tabHeadCellStyle)
          ->addText($val,
            ['size' => $options['fontSize'], 'bold' => true],
            ['alignment' => 'center']
          );
      }
    }

    foreach ($data as $row){
      $table->addRow();
      foreach($row as $c => $val) {
        $table
          ->addCell($options['width'][$c], $style['cellStyle'][$c] )
          ->addText($val, $style['fontStyle'][$c], $style['paragraphStyle'][$c]);
      }
    }

    $this->templateProcessor->setComplexBlock($block_table_name, $table);
  }

  /**
   * Save document
   *
   * @param string $file_name
   * @param string $file_path for local download, if $file_path=='' file will be downloaded
   */
  public function save(string $file_name = 'report', string $file_path = ''){
    if($file_path != ''){
      if(!is_dir($file_path)){
        echo "<h3>Не существует пути для сохранения файла: $file_path</h3>";
        return;
      }
      $this->templateProcessor->saveAs($file_path . "/" . $file_name . ".docx");
      return;
    }

    ob_start();
    $temp_file_uri = $this->templateProcessor->save();
    if(ob_get_level()==0) {
      ob_end_clean();
    }
    ob_get_clean();
    ob_clean();
    header('Content-Description: File Transfer');
    header('Content-Type: application/octet-stream');
    header('Content-Disposition: attachment; filename="' . $file_name . '.docx"');
    header('Content-Transfer-Encoding: binary');
    header('Expires: 0');
    header('Content-Length:'.filesize($temp_file_uri));
    readfile($temp_file_uri);
    unlink($temp_file_uri);
    exit;
  }
}
