<?php
/**
 * GWMD → docx 导出接口
 * POST /gwmd/api.php  body=GWMD文本
 * 返回: docx 文件下载
 */
error_reporting(0);

if ($_SERVER['REQUEST_METHOD'] !== 'POST') {
    http_response_code(405);
    header('Content-Type: application/json; charset=utf-8');
    echo json_encode(['error' => '仅支持 POST 请求'], JSON_UNESCAPED_UNICODE);
    exit;
}

$gwmd = file_get_contents('php://input');
if ($gwmd === false || strlen($gwmd) === 0) {
    http_response_code(400);
    header('Content-Type: application/json; charset=utf-8');
    echo json_encode(['error' => '请求体为空'], JSON_UNESCAPED_UNICODE);
    exit;
}
if (strlen($gwmd) > 3000000) {
    http_response_code(413);
    header('Content-Type: application/json; charset=utf-8');
    echo json_encode(['error' => '内容过大（超过3MB）'], JSON_UNESCAPED_UNICODE);
    exit;
}

$tmp_in  = tempnam(sys_get_temp_dir(), 'gwmd_');
$tmp_out = sys_get_temp_dir() . '/gwmd_out_' . uniqid() . '.docx';
file_put_contents($tmp_in, $gwmd);

$script = __DIR__ . '/gwmd2docx.py';
$cmd = 'python3 ' . escapeshellarg($script) . ' ' . escapeshellarg($tmp_in) . ' ' . escapeshellarg($tmp_out) . ' 2>&1';
exec($cmd, $out_lines, $rc);

if ($rc !== 0 || !file_exists($tmp_out)) {
    @unlink($tmp_in);
    @unlink($tmp_out);
    http_response_code(500);
    header('Content-Type: application/json; charset=utf-8');
    echo json_encode([
        'error' => 'docx 生成失败',
        'detail' => implode("\n", $out_lines)
    ], JSON_UNESCAPED_UNICODE);
    exit;
}

header('Content-Type: application/vnd.openxmlformats-officedocument.wordprocessingml.document');
header('Content-Disposition: attachment; filename="gwmd_' . date('Ymd_His') . '.docx"');
header('Content-Length: ' . filesize($tmp_out));
readfile($tmp_out);
@unlink($tmp_in);
@unlink($tmp_out);
exit;
