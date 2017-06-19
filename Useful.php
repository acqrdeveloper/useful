<?php
/**
 * Created by PhpStorm.
 * User: QuispeRoque
 * Date: 22/04/17
 * Time: 09:13
 */

namespace SF\Http\Controllers;

use Carbon\Carbon;
use Illuminate\Support\Facades\Auth;
use Illuminate\Support\Facades\DB;
use Illuminate\Support\Facades\Log;
use Intervention\Image\Facades\Image;
use PHPExcel_IOFactory;
use PHPExcel_Style_Alignment;

/**
 * variables constantes dinamicas
 */
define('FECHA_DEFAULT_FORMAT', 'Y-m-d');
define('FECHA', Carbon::now()->format(FECHA_DEFAULT_FORMAT));
define('FECHA_HORA', Carbon::now()->format(FECHA_DEFAULT_FORMAT . ' H:i:s'));
define('FECHA_DETALLE', Carbon::now()->format('Ymd') . '_' . Carbon::now()->format('His'));
define('FECHA_1MES', Carbon::now()->addMonth(1)->format(FECHA_DEFAULT_FORMAT));

trait Useful
{
    /**
     * variables Request
     */
    public $service;// lógica
    public $request;// parametros
    public $ajax;// ajax
    public $rpta = [];// respuesta generica

    /**
     * variables para PHPExcel
     */
    public $textAlignHCenter = ['alignment' => ['horizontal' => \PHPExcel_Style_Alignment::HORIZONTAL_CENTER]];
    public $textAlignVCenter = ['alignment' => ['horizontal' => \PHPExcel_Style_Alignment::VERTICAL_CENTER]];
    public $textAlignHRight = ['alignment' => ['horizontal' => \PHPExcel_Style_Alignment::HORIZONTAL_RIGHT]];
    public $textAlignHLeft = ['alignment' => ['horizontal' => \PHPExcel_Style_Alignment::HORIZONTAL_LEFT]];
    public $borderAllBordersTHIN = ['borders' => ['allborders' => ['style' => \PHPExcel_Style_Border::BORDER_THIN]]];
    public $borderOutlineTHIN = ['borders' => ['outline' => ['style' => \PHPExcel_Style_Border::BORDER_THIN]]];
    public $colorFillBlueSOLID = ['fill' => ['type' => \PHPExcel_Style_Fill::FILL_SOLID, 'color' => ['rgb' => '2196F3']]];
    public $colorFillGreenSOLID = ['fill' => ['type' => \PHPExcel_Style_Fill::FILL_SOLID, 'color' => ['rgb' => '4caf50']]];
    public $colorFillTealSOLID = ['fill' => ['type' => \PHPExcel_Style_Fill::FILL_SOLID, 'color' => ['rgb' => '009688']]];
    public $colorFillGreySOLID = ['fill' => ['type' => \PHPExcel_Style_Fill::FILL_SOLID, 'color' => ['rgb' => '9e9e9e']]];
    public $colorFillBlueGreySOLID = ['fill' => ['type' => \PHPExcel_Style_Fill::FILL_SOLID, 'color' => ['rgb' => '607d8b']]];
    public $colorFillYellowSOLID = ['fill' => ['type' => \PHPExcel_Style_Fill::FILL_SOLID, 'color' => ['rgb' => 'ffeb3b']]];
    public $colorFillAmberSOLID = ['fill' => ['type' => \PHPExcel_Style_Fill::FILL_SOLID, 'color' => ['rgb' => 'ffc107']]];
    public $colorFillOrangeSOLID = ['fill' => ['type' => \PHPExcel_Style_Fill::FILL_SOLID, 'color' => ['rgb' => 'ff9800']]];
    public $colorFillIndigoSOLID = ['fill' => ['type' => \PHPExcel_Style_Fill::FILL_SOLID, 'color' => ['rgb' => '3f51b5']]];
    public $colorFillRedSOLID = ['fill' => ['type' => \PHPExcel_Style_Fill::FILL_SOLID, 'color' => ['rgb' => 'f44336']]];
    public $colorFillNoneSOLID = ['fill' => ['type' => \PHPExcel_Style_Fill::FILL_NONE]];
    public $textDefaultBOLD = ['font' => ['bold' => true, 'color' => ['rgb' => '000000']]];
    public $textWhiteBOLD = ['font' => ['bold' => true, 'color' => ['rgb' => 'ffffff']]];
    public $textWhite = ['font' => ['bold' => false, 'color' => ['rgb' => 'ffffff']]];
    public $textBlack = ['font' => ['bold' => false, 'color' => ['rgb' => '000000']]];
    public $textGrey = ['font' => ['bold' => false, 'color' => ['rgb' => '9e9e9e']]];

    /**
     * metodo generico que intercede para el catch y prepara la respuesta generica.
     * @param $exception
     * @param string $title
     * @param string $level
     */
    function fnException($exception = null, $title = 'advertencia', $level = 'warning')
    {
        if (!is_null($exception)) {
            if ($exception->getCode() > 0) {// PDOException
                $this->rpta = ['load' => false, 'data' => null, 'message' => $exception->getPrevious()->errorInfo[2], 'title' => $title, 'level' => $level];
            } else {// Exception
                $this->rpta = ['load' => false, 'data' => null, 'message' => $exception->getMessage(), 'title' => $title, 'level' => $level];
            }
        }
    }

    /**
     * metodo generico que ingresa en la respuesta generica para notificar mensaje personalizado de Error
     * @param $error
     * @param string $title
     * @param string $message
     * @param string $level
     */
    function fnError($error = null, $title = 'ERROR', $message = 'contácte al administrador', $level = 'warning')
    {
        if (!is_null($error)) {
            $this->rpta = ['load' => false, 'data' => null, 'detail' => $error, 'title' => $title, 'message' => $message, 'level' => $level];
        }
    }

    /**
     * metodo generico que realiza la respuesta generica de satisfaccion
     * @param string $message
     * @param null $data
     * @param string $title
     * @param string $level
     */
    function fnSuccess($message = 'ejecutado correctamente', $data = null, $title = 'BIEN', $level = 'success')
    {
        $this->rpta = ['load' => true, 'data' => $data, 'title' => $title, 'message' => $message, 'level' => $level];
    }

    /**
     * metodo generico para las notificaciones a la vista.
     *
     * @param string $title
     * @param string $message
     * @param  string $level
     * @return \Laracasts\Flash\FlashNotifier
     */
    function fnFlashMessage($title = 'bien', $message = 'ejecutado correctamente', $level = 'success')
    {
        $arr_message = ['title' => $title, 'message' => $message];
        $notifier = app('flash');

        if (!is_null($message)) {
            $notifier->message($arr_message, $level);
        }
        return $notifier;
    }

    /**
     * metodo generico que genera/realiza un log segun el tipo sea indicado
     * @param $type
     * @param $detail
     */
    function fnDoLog($type, $detail)
    {
        //Establecemos zona horaria por defecto
        date_default_timezone_set('America/Lima');
        $path = storage_path() . '/logs/';

        switch ($type) {
            case 'error':
                Log::useFiles($path . 'error.log');
                Log::error($detail);
                break;
            default:
                Log::useFiles($path . 'info.log');
                Log::error($detail);
                break;
        }
    }

    /**
     * metodo generico que devuelve el maximo ID de una tabla y su AUTOINCREMENT
     * @param $table
     * @param string $field
     * @return int
     */
    function fnGetMaxID($table, $field = 'id')
    {
        $maxID = DB::table($table)->max($field);
        return (int)$maxID + 1;
    }

    /**
     * metodo generico que realiza la creacion de una hoja EXCEL
     * @param $objPHPExcel
     * @param $headers
     * @param $columns
     * @param $title
     * @param int $row
     * @param bool $merge
     * @return mixed
     */
    protected function fnCreateEXCEL($objPHPExcel, $headers, $columns, $title, $row = 1, $merge = false)
    {
        $objPHPExcel
            ->getProperties()
            ->setCreator('aquispe.developer@gmail.com')
            ->setTitle($title);

        $objPHPExcel->setActiveSheetIndex(0);
        $worksheet = $objPHPExcel->getActiveSheet();
        $total = count($columns);

        for ($i = 0; $i < $total; $i++) {
            if ($merge) {
                if (is_array($columns[$i])) {
                    $foo = $columns[$i][0] . $row . ':' . $columns[$i][1] . $row;
                    $worksheet->mergeCells($foo);
                    $worksheet->setCellValue($columns[$i][0] . $row, $headers[$i]);
                } else {
                    $worksheet->getColumnDimension($columns[$i])->setAutoSize(true);
                    $worksheet->setCellValue($columns[$i] . $row, $headers[$i]);
                }
            } else {
                $worksheet->getColumnDimension($columns[$i])->setAutoSize(true);
                $worksheet->setCellValue($columns[$i] . $row, $headers[$i]);
            }
        }

        if (is_array($columns[$total - 1])) {
            $oneColumn = $columns[$total - 1][1];
        } else {
            $oneColumn = $columns[$total - 1];
        }

        // Dejar estatico solo la primera fila
        $worksheet->freezePane('A' . ($row + 1));
        $worksheet->getStyle($columns[0] . $row . ':' . $oneColumn . $row)->applyFromArray(array($this->textDefaultBOLD, $this->textAlignHCenter));
        return $worksheet;
    }

    /**
     * metodo generico para descargar el archivo Excel
     * @param $objPHPExcel
     * @param $filename
     * @param string $type
     * @internal param null $options
     */
    protected function fnExportEXCEL($objPHPExcel, $filename, $type = 'Excel5')
    {
        $objPHPExcel->getActiveSheet()->setShowGridlines(false);
        $filename = $filename . '_' . FECHA_DETALLE;

        if ($type == 'Excel5' || $type == 'Excel2007') {
            if ($type == 'Excel5') {
                header('Content-Type: application/vnd.ms-excel');
                header('Content-Disposition: attachment;filename="' . $filename . '.xls"');
            } else {
                header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
                header('Content-Disposition: attachment;filename="' . $filename . '.xlsx"');
            }

            header('Cache-Control: max-age=0');
            // If you're serving to IE 9, then the following may be needed
            header('Cache-Control: max-age=1');

            // If you're serving to IE over SSL, then the following may be needed
            header('Expires: Mon, 26 Jul 1997 05:00:00 GMT'); // Date in the past
            header('Last-Modified: ' . gmdate('D, d M Y H:i:s') . ' GMT'); // always modified
            header('Cache-Control: cache, must-revalidate'); // HTTP/1.1
            header('Pragma: public'); // HTTP/1.0

        } else if ($type == 'PDF') {
            header('Content-Type: application/pdf');
            header('Content-Disposition: attachment="inline";filename="' . $filename . '.pdf"');
            header('Cache-Control: max-age=0');
            header("Cache-Control: private");

        }

        $excelWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, $type);
        $excelWriter->save('php://output');
    }

    /**
     * metodo generico que procesa y crea archivo pdf
     * @param $dompdf
     * @param $viewHtml
     * @param null $config
     */
    protected function fnExportPDF($dompdf, $viewHtml, $config = null)
    {
        if (is_null($config)) {
            $config = [
                'attachment' => '0',
                'hoja' => 'A4',
                'filename' => 'test.pdf',
                'orientation' => 'p',
            ];
        }
        $dompdf->loadHtml($viewHtml);
        $dompdf->setPaper($config['hoja'], $config['orientation'] == 'P' ? 'portrait' : 'landscape');
        $dompdf->render();
        $dompdf->stream($config['filename'], ['Attachment' => $config['attachment']]);
    }

    /**
     * metodo generico que devuelve una consulta SELECT con/sin INNER JOIN para el autocomplete de boostrap
     * @param $table
     * @param string $columns
     * @param null $colum_state
     * @param bool $flag_join
     * @param null $tablejoin
     * @param null $tableid
     * @param null $tablejoinid
     * @return \Illuminate\Http\JsonResponse
     */
    protected function fnGetListAutocomplete($table, $columns = '*', $colum_state = null, $flag_join = false, $tablejoin = null, $tableid = null, $tablejoinid = null)
    {
        try {
            if ($flag_join) {
                if (!is_null($colum_state)) {
                    $data = DB::table($table)->join($tablejoin, $tableid, '=', $tablejoinid)->select($columns)->where($colum_state, 'A')->get();
                } else {
                    $data = DB::table($table)->join($tablejoin, $tableid, '=', $tablejoinid)->select($columns)->get();
                }
            } else {
                if (!is_null($colum_state)) {
                    $data = DB::table($table)->select($columns)->where($colum_state, 'A')->get();
                } else {
                    $data = DB::table($table)->select($columns)->get();
                }
            }
            $arr = [];
            foreach ($data as $key => $value) {
                array_push($arr, ['value' => strtoupper($value->nombre_completo . ' - ' . $value->nombre), 'data' => $value]);
            }
            return response()->json($arr, 200);
        } catch (\PDOException $e) {
            return response()->json($e->getMessage(), 412);
        }
    }

    /**
     * metodo generico que almacena la imagen en server
     * @param $path
     * @param $request
     */
    protected function fnSaveImage($path, $request)
    {
        $image = Image::make($request);
        $image->save($path . $request->getClientOriginalName());
    }

}