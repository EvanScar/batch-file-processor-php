<?php
session_start();
require_once __DIR__ . '/../vendor/autoload.php';

use PhpOffice\PhpWord\PhpWord;
use PhpOffice\PhpWord\IOFactory as WordIOFactory;
use PhpOffice\PhpSpreadsheet\IOFactory as SpreadsheetIOFactory;

error_reporting(0);
ini_set('display_errors', 0);
header('Content-Type: application/json; charset=utf-8');

try {
    if ($_SERVER['REQUEST_METHOD'] !== 'POST') throw new Exception('Только POST');
    if (!isset($_FILES['files']) || empty($_FILES['files']['name'][0])) throw new Exception('Файлы не выбраны');

    $op = $_POST['operation'] ?? 'rename';
    $param = $_POST['pattern'] ?? $_POST['format'] ?? '';
    
    $tempDir = sys_get_temp_dir() . '/fp_' . session_id();
    $outDir = $tempDir . '/out/';
    if (!is_dir($tempDir)) mkdir($tempDir, 0777, true);
    if (!is_dir($outDir)) mkdir($outDir, 0777, true);

    array_map('unlink', glob("$outDir*"));

    $files = $_FILES['files'];
    $names = is_array($files['name']) ? $files['name'] : [$files['name']];
    $tmps = is_array($files['tmp_name']) ? $files['tmp_name'] : [$files['tmp_name']];
    $errs = is_array($files['error']) ? $files['error'] : [$files['error']];

    $results = [];
    $errors = [];
    $pdfImages = [];

    foreach ($names as $i => $name) {
        if ($errs[$i] !== UPLOAD_ERR_OK) { $errors[] = "Ошибка загрузки: $name"; continue; }
        $tmp = $tmps[$i];
        $ext = strtolower(pathinfo($name, PATHINFO_EXTENSION));
        $base = pathinfo($name, PATHINFO_FILENAME);

        try {
            if ($op === 'rename') {
                $new = "$param_" . str_pad($i+1, 3, '0', STR_PAD_LEFT) . ".$ext";
                move_uploaded_file($tmp, "$outDir$new");
                $results[] = ['name'=>$new, 'path'=>"$outDir$new"];
            } 
            elseif ($op === 'img_convert') {
                $imgInfo = @getimagesize($tmp);
                if (!$imgInfo) { $errors[] = "$name: не изображение"; continue; }
                
                $pdfReadyPath = $tmp;
                if (in_array($ext, ['webp', 'bmp'])) {
                    $pdfReadyPath = "$tempDir/conv_{$i}.jpg";
                    $imgRes = imagecreatefromstring(file_get_contents($tmp));
                    imagejpeg($imgRes, $pdfReadyPath, 90);
                    imagedestroy($imgRes);
                }

                if ($param === 'pdf') {
                    $pdfImages[] = $pdfReadyPath;
                    continue;
                }

                $imgRes = imagecreatefromstring(file_get_contents($tmp));
                $out = "$outDir$base.$param";
                
                if ($param === 'jpg') imagejpeg($imgRes, $out, 90);
                elseif ($param === 'png') imagepng($imgRes, $out);
                elseif ($param === 'gif') imagegif($imgRes, $out);
                elseif ($param === 'webp') imagewebp($imgRes, $out, 80);
                else throw new Exception("Формат не поддерживается");
                
                imagedestroy($imgRes);
                $results[] = ['name'=>"$base.$param", 'path'=>$out];
            }
            elseif ($op === 'doc_convert') {
                $rawText = '';
                $tableData = [];
                $hasTable = false;

                // === НАДЁЖНОЕ ИЗВЛЕЧЕНИЕ ТЕКСТА ===
                if (in_array($ext, ['xlsx', 'xls', 'csv'])) {
                    $spreadsheet = SpreadsheetIOFactory::load($tmp);
                    $sheet = $spreadsheet->getActiveSheet();
                    foreach ($sheet->getRowIterator() as $row) {
                        $cells = [];
                        foreach ($row->getCellIterator() as $cell) {
                            $cells[] = $cell->getFormattedValue();
                        }
                        $tableData[] = $cells;
                        $rawText .= implode(" | ", $cells) . "\n";
                    }
                    $hasTable = !empty($tableData);
                } 
                elseif ($ext === 'docx') {
                    $phpWord = WordIOFactory::load($tmp);
                    foreach ($phpWord->getSections() as $section) {
                        foreach ($section->getElements() as $el) {
                            if ($el instanceof \PhpOffice\PhpWord\Element\Table) {
                                foreach ($el->getRows() as $row) {
                                    $rowData = [];
                                    foreach ($row->getCells() as $cell) {
                                        $cellText = '';
                                        foreach ($cell->getElements() as $cellEl) {
                                            if ($cellEl instanceof \PhpOffice\PhpWord\Element\TextRun) {
                                                foreach ($cellEl->getElements() as $t) {
                                                    if (method_exists($t, 'getText')) $cellText .= $t->getText();
                                                }
                                            } elseif (method_exists($cellEl, 'getText')) {
                                                $cellText .= $cellEl->getText();
                                            }
                                        }
                                        $rowData[] = trim($cellText);
                                    }
                                    $tableData[] = $rowData;
                                }
                                $hasTable = true;
                            } elseif ($el instanceof \PhpOffice\PhpWord\Element\TextRun) {
                                foreach ($el->getElements() as $t) {
                                    if (method_exists($t, 'getText')) $rawText .= $t->getText() . " ";
                                }
                                $rawText .= "\n";
                            } elseif (method_exists($el, 'getText')) {
                                $rawText .= $el->getText() . "\n";
                            }
                        }
                    }
                } 
                elseif ($ext === 'xml') {
                    $dom = new DOMDocument();
                    $dom->preserveWhiteSpace = false;
                    if (@$dom->load($tmp)) {
                        $rawText = $dom->textContent;
                    } else {
                        $rawText = strip_tags(file_get_contents($tmp));
                    }
                } 
                elseif ($ext === 'txt') {
                    $rawText = file_get_contents($tmp);
                } 
                else {
                    throw new Exception("Формат не поддерживается");
                }

                // Нормализация и принудительный UTF-8
                $rawText = preg_replace('/\s+/', ' ', $rawText);
                $rawText = str_replace(["\r\n", "\r"], "\n", $rawText);
                if (!mb_check_encoding($rawText, 'UTF-8')) {
                    $rawText = mb_convert_encoding($rawText, 'UTF-8', 'auto');
                }
                $rawText = trim($rawText);
                $out = "$outDir$base.$param";

                // === ГЕНЕРАЦИЯ ВЫХОДНОГО ФАЙЛА ===
                if ($param === 'pdf') {
                    $pdf = new TCPDF();
                    $pdf->SetFont('dejavusans', '', 10); // Встроенный шрифт TCPDF с кириллицей
                    $pdf->AddPage();
                    
                    if ($hasTable && !empty($tableData)) {
                        $html = '<table border="1" cellpadding="4" cellspacing="0" style="font-size:10px;">';
                        foreach ($tableData as $rIdx => $row) {
                            $html .= '<tr>';
                            $tag = ($rIdx === 0) ? 'th' : 'td';
                            $style = ($rIdx === 0) ? ' style="background:#eee; font-weight:bold;"' : '';
                            foreach ($row as $cell) {
                                $html .= "<$tag$style>" . htmlspecialchars($cell, ENT_QUOTES, 'UTF-8') . "</$tag>";
                            }
                            $html .= '</tr>';
                        }
                        $html .= '</table>';
                        $pdf->writeHTML($html, true, false, true, false, '');
                    } else {
                        $pdf->MultiCell(0, 6, $rawText, false, 'L', false, 1, '', '', true, 0, false, true, 0, 'T', 'J');
                    }
                    $pdf->Output($out, 'F');
                } 
                elseif ($param === 'docx') {
                    $phpWord = new PhpWord();
                    $phpWord->setDefaultFontName('Arial');
                    $phpWord->setDefaultFontSize(11);
                    $section = $phpWord->addSection();
                    
                    if ($hasTable && !empty($tableData)) {
                        $table = $section->addTable(['borderSize' => 6, 'borderColor' => '999999', 'cellMargin' => 60]);
                        foreach ($tableData as $rIdx => $row) {
                            $rowStyle = ($rIdx === 0) ? ['bgColor' => 'EEEEEE'] : null;
                            $tableRow = $table->addRow(300, $rowStyle);
                            foreach ($row as $cellText) {
                                $cell = $tableRow->addCell(2000);
                                if ($rIdx === 0) {
                                    $cell->addText($cellText, ['bold' => true, 'name' => 'Arial']);
                                } else {
                                    $cell->addText($cellText, ['name' => 'Arial']);
                                }
                            }
                        }
                    } else {
                        $section->addText($rawText, ['name' => 'Arial']);
                    }
                    WordIOFactory::createWriter($phpWord, 'Word2007')->save($out);
                } 
                elseif ($param === 'txt') {
                    file_put_contents($out, $rawText);
                } 
                elseif ($param === 'csv') {
                    if ($hasTable && !empty($tableData)) { 
                        $fp = fopen($out, 'w'); 
                        foreach($tableData as $r) fputcsv($fp, $r); 
                        fclose($fp); 
                    } else throw new Exception("CSV требует табличных данных");
                }
                elseif ($param === 'xml') {
                    $xml = new SimpleXMLElement('<document/>');
                    $xml->addAttribute('converted_from', strtoupper($ext));
                    $xml->addAttribute('generated_at', date('Y-m-d H:i:s'));
                    $xml->addChild('metadata', "Converted by File Processor");
                    if ($hasTable && !empty($tableData)) {
                        $tNode = $xml->addChild('table');
                        foreach ($tableData as $rIdx => $row) {
                            $rNode = $tNode->addChild('row')->addAttribute('type', $rIdx===0?'header':'data');
                            foreach ($row as $cIdx => $cell) {
                                $rNode->addChild('cell', htmlspecialchars($cell, ENT_XML1, 'UTF-8'))->addAttribute('index', $cIdx);
                            }
                        }
                    } else {
                        $xml->addChild('content')->addChild('text', htmlspecialchars($rawText, ENT_XML1, 'UTF-8'));
                    }
                    $xml->asXML($out);
                }
                else {
                    throw new Exception("Целевой формат не поддерживается");
                }

                $results[] = ['name'=>"$base.$param", 'path'=>$out];
            }
        } catch (Exception $e) {
            $errors[] = "$name: " . $e->getMessage();
        }
    }

    if ($op === 'img_convert' && $param === 'pdf' && !empty($pdfImages)) {
        $pdf = new TCPDF();
        $pdf->SetFont('dejavusans', '', 12);
        foreach ($pdfImages as $idx => $pImg) {
            $pdf->AddPage();
            $info = getimagesize($pImg);
            $scale = min(190/$info[0], 277/$info[1]);
            $pdf->Image($pImg, (210-$info[0]*$scale)/2, (297-$info[1]*$scale)/2, $info[0]*$scale, $info[1]*$scale);
        }
        $pdfOut = "$outDir/documents.pdf";
        $pdf->Output($pdfOut, 'F');
        $results[] = ['name'=>'documents.pdf', 'path'=>$pdfOut];
    }

    $_SESSION['tokens'] = [];
    $tokens = [];
    foreach ($results as $r) {
        $t = bin2hex(random_bytes(12));
        $_SESSION['tokens'][$t] = ['path'=>$r['path'], 'name'=>$r['name'], 'size'=>filesize($r['path']), 'exp'=>time()+3600];
        $tokens[] = ['token'=>$t, 'name'=>$r['name'], 'size'=>filesize($r['path'])];
    }

    $zipToken = null;
    if (count($results) > 1) {
        $zipPath = "$outDir/result.zip";
        $zip = new ZipArchive();
        if ($zip->open($zipPath, ZipArchive::CREATE | ZipArchive::OVERWRITE) === true) {
            foreach ($results as $r) $zip->addFile($r['path'], $r['name']);
            $zip->close();
            $zt = bin2hex(random_bytes(12));
            $_SESSION['tokens'][$zt] = ['path'=>$zipPath, 'name'=>'result.zip', 'size'=>filesize($zipPath), 'exp'=>time()+3600];
            $zipToken = $zt;
        }
    }

    echo json_encode(['status'=>'ok', 'files'=>$tokens, 'zipToken'=>$zipToken, 'count'=>count($results), 'errors'=>$errors], JSON_UNESCAPED_UNICODE);

} catch (Exception $e) {
    http_response_code(500);
    echo json_encode(['status'=>'error', 'message'=>$e->getMessage()], JSON_UNESCAPED_UNICODE);
}"// TODO: Add JSON conversion logic" 
/ /   T O D O :   A d d   J S O N   c o n v e r s i o n   l o g i c  
 