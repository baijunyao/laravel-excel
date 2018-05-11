<?php

if (! function_exists('export_excel')) {
    /**
     * 导出excel文件
     *
     * @param $data
     * @param string $file_name
     * @param string $ext
     *
     *  单个工作表 示例数组：
     *  $data = array(
     *      ['data1', 'data2'],
     *      ['data5', 'data6'],
     *      ['data3', 'data4']
     *  );
     *
     *  多个工作表 示例数组：
     *  $data = array(
     *      'Sheet1' => ['data1', 'data2'],
     *      'Sheet2' => ['data5', 'data6'],
     *      'Sheet3' => ['data3', 'data4']
     *  );
     *
     */
    function export_excel($data, $file_name = 'filename', $ext = 'xls')
    {
        foreach ($data as $k => $v) {
            // 如果是数组 直接返回
            if (is_array($v)) {
                continue;
            }
            // 如果是 stdClass 则转成数据
            if (is_object($v) && in_array(get_class($v), ['stdClass'])) {
                $data[$k] = (array)$v;
            }
        }

        // 利用array_merge把键设置连续数组；为后续判断是否为关联数组做准备
        $data = array_merge($data);

        // 如果是索引数组 则直接创建单个工作表的excel
        if (array_values($data) === $data) {
            Excel::create($file_name, function($excel) use($data) {
                $excel->sheet('Sheet1', function($sheet) use($data)  {
                    $sheet->fromArray($data, null, 'A1', false, false);
                });
            })->export($ext);
        } else {
            // 如果是关联数组；则生成 名字为key的工作表sheet 的excel
            Excel::create($file_name, function($excel) use($data) {
                foreach ($data as $k => $v) {
                    $excel->sheet($k, function($sheet) use($v)  {
                        $sheet->fromArray($v, null, 'A1', false, false);
                    });
                }
            })->export($ext);
        }
    }
}

if (! function_exists('import_excel')) {
    /**
     * 导入excel文件
     *
     * @param       $file
     * @param array $replace  示例 [null => '']  则会把所有值为 null 的替换为'' 空字符串
     *
     * @return mixed
     */
    function import_excel($file, $replace = [])
    {
        $excel = Excel::load($file)->get();
        if (!empty($replace)) {
            $excel = $excel->map(function ($v) use ($replace) {
                $search = array_keys($replace);
                foreach ($v as $m => $n) {
                    if (in_array($n, $search)) {
                        $v[$m] = $replace[$n];
                    }
                }
                return $v;
            });
        }
        return $excel;
    }
}