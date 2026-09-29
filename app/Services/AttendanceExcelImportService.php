<?php

namespace App\Services;

use App\Models\Employee;
use App\Models\EmployeeAttendance;
use App\Models\Setting;
use DateTime;
use Illuminate\Support\Facades\DB;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Shared\Date as ExcelDate;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use Throwable;

class AttendanceExcelImportService
{
    /**
     * Import attendance spreadsheet file (.xlsx, .xls, .csv).
     *
     * @return array{employees: int, total_employees: int, attendance: int, skipped: int, sheets: int, errors: array}
     */
    public function import(string $filePath): array
    {
        $spreadsheet = IOFactory::load($filePath);

        $result = [
            'employees' => 0,
            'total_employees' => 0,
            'attendance' => 0,
            'skipped' => 0,
            'sheets' => 0,
            'errors' => [],
        ];

        $affectedEmployeeIds = [];
        $setting = Setting::first();

        DB::beginTransaction();

        try {
            foreach ($spreadsheet->getWorksheetIterator() as $sheet) {
                $result['sheets']++;

                try {
                    if ($this->looksLikeFingerprintRecord($sheet)) {
                        $this->importFingerprintSheet($sheet, $result, $affectedEmployeeIds);
                    } elseif ($this->looksLikeHorizontalAttendance($sheet)) {
                        $this->importHorizontalSheet($sheet, $result, $affectedEmployeeIds, $setting);
                    } elseif ($this->looksLikeMatrixAttendance($sheet)) {
                        $this->importMatrixSheet($sheet, $result, $affectedEmployeeIds);
                    } else {
                        $this->importVerticalSheet($sheet, $result, $affectedEmployeeIds);
                    }
                } catch (Throwable $e) {
                    $result['errors'][] = [
                        'sheet' => $sheet->getTitle(),
                        'message' => $e->getMessage(),
                    ];
                }
            }

            // Synchronize Employee summary models with their imported attendances
            $this->syncEmployees(array_unique($affectedEmployeeIds));

            DB::commit();
        } catch (Throwable $e) {
            DB::rollBack();
            throw $e;
        }

        $result['total_employees'] = Employee::count();

        return $result;
    }

    /**
     * Detects fingerprint attendance machine format (e.g. ZKTeco, Solution, Fingerspot).
     */
    private function looksLikeFingerprintRecord(Worksheet $sheet): bool
    {
        $maxRow = min(20, (int) $sheet->getHighestRow());

        for ($r = 1; $r <= $maxRow; $r++) {
            for ($c = 1; $c <= 15; $c++) {
                $col = Coordinate::stringFromColumnIndex($c);
                $val = (string) $sheet->getCell("{$col}{$r}")->getValue();

                if (
                    stripos($val, 'Employee Attendance Record') !== false ||
                    stripos($val, 'User ID:') !== false
                ) {
                    return true;
                }
            }
        }

        return false;
    }

    /**
     * Imports fingerprint attendance machine format.
     */
    private function importFingerprintSheet(Worksheet $sheet, array &$result, array &$affectedEmployeeIds): void
    {
        $highestRow = (int) $sheet->getHighestRow();

        // 1. Detect date range (e.g. "Attendance date:2026-08-01~2026-08-30")
        $year = (int) date('Y');
        $month = (int) date('m');

        for ($r = 1; $r <= min(10, $highestRow); $r++) {
            for ($c = 1; $c <= 30; $c++) {
                $col = Coordinate::stringFromColumnIndex($c);
                $val = (string) $sheet->getCell("{$col}{$r}")->getValue();

                if (stripos($val, 'Attendance date:') !== false) {
                    $dateRangeStr = trim(str_ireplace('Attendance date:', '', $val));
                    $parts = explode('~', $dateRangeStr);
                    if (! empty($parts[0])) {
                        $parsedDate = strtotime(trim($parts[0]));
                        if ($parsedDate !== false) {
                            $year = (int) date('Y', $parsedDate);
                            $month = (int) date('m', $parsedDate);
                            break 2;
                        }
                    }
                }
            }
        }

        // 2. Scan all employees in this sheet
        $users = [];
        for ($r = 1; $r <= $highestRow; $r++) {
            for ($c = 1; $c <= 12; $c++) {
                $col = Coordinate::stringFromColumnIndex($c);
                $val = trim((string) $sheet->getCell("{$col}{$r}")->getValue());

                if (strcasecmp($val, 'User ID:') === 0) {
                    $idCol = Coordinate::stringFromColumnIndex($c + 1);
                    $userId = trim((string) $sheet->getCell("{$idCol}{$r}")->getValue());

                    $name = '';
                    for ($nc = $c + 2; $nc <= $c + 12; $nc++) {
                        $colN = Coordinate::stringFromColumnIndex($nc);
                        if (strcasecmp(trim((string) $sheet->getCell("{$colN}{$r}")->getValue()), 'Name:') === 0) {
                            $valCol = Coordinate::stringFromColumnIndex($nc + 1);
                            $name = trim((string) $sheet->getCell("{$valCol}{$r}")->getValue());
                            break;
                        }
                    }

                    $dept = 'COMPANY';
                    for ($dc = $c + 8; $dc <= 30; $dc++) {
                        $colD = Coordinate::stringFromColumnIndex($dc);
                        if (strcasecmp(trim((string) $sheet->getCell("{$colD}{$r}")->getValue()), 'Department:') === 0) {
                            $deptValCol = Coordinate::stringFromColumnIndex($dc + 1);
                            $dept = trim((string) $sheet->getCell("{$deptValCol}{$r}")->getValue());
                            break;
                        }
                    }

                    if ($name !== '') {
                        $users[] = [
                            'row' => $r,
                            'id' => $userId,
                            'name' => $name,
                            'dept' => $dept,
                        ];
                    }
                    break;
                }
            }
        }

        if (empty($users)) {
            $result['skipped']++;

            return;
        }

        $userCount = count($users);

        for ($u = 0; $u < $userCount; $u++) {
            $currentUser = $users[$u];
            $nextUserRow = ($u + 1 < $userCount) ? $users[$u + 1]['row'] : $highestRow + 1;
            $userRow = $currentUser['row'];

            // Find day numbers in row directly following userRow
            $dayRow = $userRow + 1;
            $dayCols = [];

            for ($c = 2; $c <= 35; $c++) {
                $col = Coordinate::stringFromColumnIndex($c);
                $val = trim((string) $sheet->getCell("{$col}{$dayRow}")->getValue());
                if (is_numeric($val) && (int) $val >= 1 && (int) $val <= 31) {
                    $dayCols[(int) $val] = $col;
                }
            }

            if (empty($dayCols)) {
                continue;
            }

            $formattedId = is_numeric($currentUser['id'])
                ? 'EMP-'.str_pad($currentUser['id'], 3, '0', STR_PAD_LEFT)
                : 'EMP-'.$currentUser['id'];

            $employee = $this->findOrCreateEmployee(
                $currentUser['name'],
                $formattedId,
                $currentUser['dept']
            );

            if ($employee->wasRecentlyCreated) {
                $result['employees']++;
            }

            $affectedEmployeeIds[] = $employee->id;

            // Extract punch times from rows between $dayRow + 1 and $nextUserRow - 1
            foreach ($dayCols as $dayNum => $colLetter) {
                $dateStr = sprintf('%04d-%02d-%02d', $year, $month, $dayNum);
                $punchStrings = [];

                for ($pr = $dayRow + 1; $pr < $nextUserRow; $pr++) {
                    $cellVal = (string) $sheet->getCell("{$colLetter}{$pr}")->getValue();
                    if (trim($cellVal) !== '') {
                        foreach (explode("\n", $cellVal) as $line) {
                            $cleanLine = trim($line);
                            if ($cleanLine !== '') {
                                $punchStrings[] = $cleanLine;
                            }
                        }
                    }
                }

                $timeIn = null;
                $timeOut = null;
                $note = null;

                if (! empty($punchStrings)) {
                    $validTimes = [];
                    foreach ($punchStrings as $ps) {
                        $parsed = $this->parseTime($ps);
                        if ($parsed !== null) {
                            $validTimes[] = $parsed;
                        } else {
                            $note = $note ? $note.' | '.$ps : $ps;
                        }
                    }

                    if (! empty($validTimes)) {
                        sort($validTimes);
                        $timeIn = $validTimes[0];
                        if (count($validTimes) > 1) {
                            $timeOut = $validTimes[count($validTimes) - 1];
                        }
                    }
                }

                $statusData = $this->detectStatusAndDeduction($timeIn, $timeOut, $note, $dateStr);

                EmployeeAttendance::updateOrCreate(
                    [
                        'employee_id' => $employee->id,
                        'date' => $dateStr,
                    ],
                    [
                        'time_in' => $timeIn,
                        'time_out' => $timeOut,
                        'status' => $statusData['status'],
                        'deduction' => $statusData['deduction'],
                        'note' => $statusData['note'],
                    ]
                );

                $result['attendance']++;
            }
        }
    }

    /**
     * Detects horizontal multi-employee IN/OUT format (e.g. Maintenance attendance sheets).
     */
    private function looksLikeHorizontalAttendance(Worksheet $sheet): bool
    {
        $rows = $sheet->toArray(null, true, false, true);
        $maxRows = min(20, count($rows));

        for ($r = 1; $r <= $maxRows; $r++) {
            $currentRow = $rows[$r] ?? [];
            $hasTglOrName = false;

            foreach ($currentRow as $value) {
                $text = $this->normalizeText($value);
                if (in_array($text, ['tgl', 'tanggal', 'date', 'nama', 'karyawan', 'name'], true)) {
                    $hasTglOrName = true;
                    break;
                }
            }

            if ($hasTglOrName) {
                $nextMax = min($r + 5, $maxRows);
                for ($next = $r + 1; $next <= $nextMax; $next++) {
                    $nextRow = $rows[$next] ?? [];
                    $inCount = 0;
                    $outCount = 0;

                    foreach ($nextRow as $val) {
                        $norm = $this->normalizeText($val);
                        if ($norm === 'in') {
                            $inCount++;
                        }
                        if ($norm === 'out') {
                            $outCount++;
                        }
                    }

                    if ($inCount >= 1 && $outCount >= 1) {
                        return true;
                    }
                }
            }
        }

        return false;
    }

    /**
     * Imports horizontal multi-employee format with paired IN/OUT columns.
     */
    private function importHorizontalSheet(
        Worksheet $sheet,
        array &$result,
        array &$affectedEmployeeIds,
        ?Setting $setting
    ): void {
        $rows = $sheet->toArray(null, true, false, true);

        $nameRow = null;
        $ioRow = null;
        $dateColumn = null;

        foreach (array_slice($rows, 0, 15, true) as $rowNumber => $row) {
            foreach ($row as $column => $value) {
                $text = $this->normalizeText($value);
                if (in_array($text, ['tgl', 'tanggal', 'date', 'ket'], true)) {
                    $dateColumn = $column;
                    $nameRow = $rowNumber;
                    break 2;
                }
            }
        }

        if ($dateColumn === null || $nameRow === null) {
            $result['skipped']++;

            return;
        }

        for ($r = $nameRow + 1; $r <= $nameRow + 5; $r++) {
            $row = $rows[$r] ?? [];
            $inOutFound = 0;

            foreach ($row as $value) {
                $text = $this->normalizeText($value);
                if ($text === 'in' || $text === 'out') {
                    $inOutFound++;
                }
            }

            if ($inOutFound >= 2) {
                $ioRow = $r;
                break;
            }
        }

        if ($ioRow === null) {
            $result['skipped']++;

            return;
        }

        // Detect Year and Month
        $month = null;
        $year = null;

        for ($r = max(1, $nameRow - 2); $r <= $ioRow; $r++) {
            foreach (($rows[$r] ?? []) as $value) {
                $text = $this->normalizeText($value);

                $detectedMonth = $this->monthNumber($text);
                if ($detectedMonth !== null) {
                    $month = $detectedMonth;
                }

                if (preg_match('/^(19|20)\d{2}$/', $text)) {
                    $year = (int) $text;
                } elseif (preg_match('/^\b(\d{2})\b$/', $text) && (int) $text >= 20 && (int) $text <= 40) {
                    $year = 2000 + (int) $text;
                }
            }
        }

        // Fallback year/month from sheet title
        if ($year === null || $month === null) {
            $titleNorm = $this->normalizeText($sheet->getTitle());
            if ($month === null) {
                foreach (['januari', 'februari', 'maret', 'april', 'mei', 'juni', 'juli', 'agustus', 'september', 'oktober', 'november', 'desember', 'jan', 'feb', 'mrt', 'apr', 'mei', 'jun', 'jul', 'agt', 'agu', 'sep', 'okt', 'nov', 'des'] as $m) {
                    if (str_contains($titleNorm, $m)) {
                        $month = $this->monthNumber($m);
                        break;
                    }
                }
            }

            if (preg_match('/(20\d{2})/', $titleNorm, $mYear)) {
                $year = (int) $mYear[1];
            } elseif (preg_match('/\b(\d{2})\b/', $titleNorm, $mYear) && (int) $mYear[1] >= 20) {
                $year = 2000 + (int) $mYear[1];
            }
        }

        $year ??= (int) date('Y');
        $month ??= (int) date('m');

        // Pair IN / OUT columns for each employee
        $employees = [];
        $lastSeenName = null;

        foreach (($rows[$nameRow] ?? []) as $col => $val) {
            $nameCandidate = trim((string) ($val ?? ''));
            if ($nameCandidate !== '' && ! $this->isNoiseName($nameCandidate)) {
                $lastSeenName = $nameCandidate;
            }

            // Check row above or row below if name is there
            if (! $lastSeenName && isset($rows[$nameRow + 1][$col])) {
                $subCandidate = trim((string) $rows[$nameRow + 1][$col]);
                if ($subCandidate !== '' && ! $this->isNoiseName($subCandidate) && $this->normalizeText($subCandidate) !== 'in' && $this->normalizeText($subCandidate) !== 'out') {
                    $lastSeenName = $subCandidate;
                }
            }

            $io = $this->normalizeText($rows[$ioRow][$col] ?? '');
            if ($io === 'in' || $io === 'out') {
                $colIndex = Coordinate::columnIndexFromString($col);
                // Group pairs (cols 2 & 3, 4 & 5, etc.)
                $pairGroup = intdiv($colIndex, 2);

                if (! isset($employees[$pairGroup])) {
                    $employees[$pairGroup] = [
                        'name' => $lastSeenName ?? "Karyawan {$pairGroup}",
                        'in_col' => null,
                        'out_col' => null,
                    ];
                }

                if ($lastSeenName && empty($employees[$pairGroup]['name'])) {
                    $employees[$pairGroup]['name'] = $lastSeenName;
                }

                if ($io === 'in') {
                    $employees[$pairGroup]['in_col'] = $col;
                } else {
                    $employees[$pairGroup]['out_col'] = $col;
                }
            }
        }

        // Import daily attendance rows
        $totalRows = count($rows);
        $dept = 'MAINTENANCE';

        for ($r = $ioRow + 1; $r <= $totalRows; $r++) {
            $dateVal = $rows[$r][$dateColumn] ?? null;
            $day = $this->extractDay($dateVal);

            if ($day === null || $day < 1 || $day > 31) {
                continue;
            }

            $dateStr = sprintf('%04d-%02d-%02d', $year, $month, $day);

            foreach ($employees as $empData) {
                $empName = trim((string) ($empData['name'] ?? ''));
                if ($empName === '' || $this->isNoiseName($empName)) {
                    continue;
                }

                $timeIn = null;
                $timeOut = null;
                $note = null;

                if (! empty($empData['in_col'])) {
                    $rawIn = $rows[$r][$empData['in_col']] ?? null;
                    $parsedIn = $this->parseTime($rawIn);
                    if ($parsedIn !== null) {
                        $timeIn = $parsedIn;
                    } elseif ($rawIn !== null && trim((string) $rawIn) !== '') {
                        $note = trim((string) $rawIn);
                    }
                }

                if (! empty($empData['out_col'])) {
                    $rawOut = $rows[$r][$empData['out_col']] ?? null;
                    $parsedOut = $this->parseTime($rawOut);
                    if ($parsedOut !== null) {
                        $timeOut = $parsedOut;
                    } elseif ($rawOut !== null && trim((string) $rawOut) !== '') {
                        $note = $note ? $note.' | '.trim((string) $rawOut) : trim((string) $rawOut);
                    }
                }

                $employeeModel = $this->findOrCreateEmployee($empName, null, $dept);

                if ($employeeModel->wasRecentlyCreated) {
                    $result['employees']++;
                }

                $affectedEmployeeIds[] = $employeeModel->id;

                $statusData = $this->detectStatusAndDeduction($timeIn, $timeOut, $note, $dateStr);

                EmployeeAttendance::updateOrCreate(
                    [
                        'employee_id' => $employeeModel->id,
                        'date' => $dateStr,
                    ],
                    [
                        'time_in' => $timeIn,
                        'time_out' => $timeOut,
                        'status' => $statusData['status'],
                        'deduction' => $statusData['deduction'],
                        'note' => $statusData['note'],
                        'shift_name' => $setting?->shift_name ?? 'Shift 1',
                        'shift_start' => $setting?->start_time,
                        'shift_end' => $setting?->end_time,
                    ]
                );

                $result['attendance']++;
            }
        }
    }

    /**
     * Detects matrix attendance format (e.g. Row with NO, NAMA, 1, 2, 3... 31).
     */
    private function looksLikeMatrixAttendance(Worksheet $sheet): bool
    {
        $maxRow = min(15, (int) $sheet->getHighestRow());

        for ($r = 1; $r <= $maxRow; $r++) {
            $hasName = false;
            $numericDaysCount = 0;

            for ($c = 1; $c <= 35; $c++) {
                $col = Coordinate::stringFromColumnIndex($c);
                $val = trim((string) $sheet->getCell("{$col}{$r}")->getValue());
                $norm = $this->normalizeText($val);

                if (in_array($norm, ['nama', 'name', 'karyawan'], true)) {
                    $hasName = true;
                }

                if (is_numeric($val) && (int) $val >= 1 && (int) $val <= 31) {
                    $numericDaysCount++;
                }
            }

            if ($hasName && $numericDaysCount >= 5) {
                return true;
            }
        }

        return false;
    }

    /**
     * Imports matrix attendance format.
     */
    private function importMatrixSheet(Worksheet $sheet, array &$result, array &$affectedEmployeeIds): void
    {
        $highestRow = (int) $sheet->getHighestRow();
        $rows = $sheet->toArray(null, true, false, true);

        $headerRow = null;
        $nameCol = null;
        $dayCols = [];

        // Find header row with days 1..31
        for ($r = 1; $r <= min(15, count($rows)); $r++) {
            $row = $rows[$r] ?? [];
            $hasName = false;
            $days = [];

            foreach ($row as $c => $val) {
                $norm = $this->normalizeText($val);
                if (in_array($norm, ['nama', 'name', 'karyawan'], true)) {
                    $nameCol = $c;
                    $hasName = true;
                }

                if (is_numeric($val) && (int) $val >= 1 && (int) $val <= 31) {
                    $days[(int) $val] = $c;
                }
            }

            if ($hasName && count($days) >= 5) {
                $headerRow = $r;
                $dayCols = $days;
                break;
            }
        }

        if ($headerRow === null || $nameCol === null || empty($dayCols)) {
            $result['skipped']++;

            return;
        }

        // Detect Year and Month
        $year = (int) date('Y');
        $month = (int) date('m');

        for ($r = 1; $r <= $headerRow; $r++) {
            foreach (($rows[$r] ?? []) as $val) {
                $text = $this->normalizeText($val);
                $detMonth = $this->monthNumber($text);
                if ($detMonth !== null) {
                    $month = $detMonth;
                }

                if (preg_match('/(20\d{2})/', $text, $m)) {
                    $year = (int) $m[1];
                }
            }
        }

        $sheetTitleNorm = $this->normalizeText($sheet->getTitle());
        $titleMonth = $this->monthNumber($sheetTitleNorm);
        if ($titleMonth !== null) {
            $month = $titleMonth;
        }
        if (preg_match('/(20\d{2})/', $sheetTitleNorm, $tm)) {
            $year = (int) $tm[1];
        }

        $dept = 'MAINTENANCE';

        for ($r = $headerRow + 1; $r <= count($rows); $r++) {
            $name = trim((string) ($rows[$r][$nameCol] ?? ''));
            if ($name === '' || $this->isNoiseName($name)) {
                continue;
            }

            $employee = $this->findOrCreateEmployee($name, null, $dept);
            if ($employee->wasRecentlyCreated) {
                $result['employees']++;
            }
            $affectedEmployeeIds[] = $employee->id;

            foreach ($dayCols as $dayNum => $col) {
                $val = trim((string) ($rows[$r][$col] ?? ''));
                if ($val === '') {
                    continue;
                }

                $dateStr = sprintf('%04d-%02d-%02d', $year, $month, $dayNum);
                $timeIn = null;
                $timeOut = null;
                $note = null;
                $status = 'hadir';
                $deduction = 0;

                if ($val === '1') {
                    $status = 'hadir';
                    $timeIn = '08:00:00';
                    $timeOut = '17:00:00';
                } elseif ($val === '1/2' || $val === '0.5') {
                    $status = 'telat';
                    $timeIn = '08:30:00';
                    $timeOut = '17:00:00';
                    $deduction = 50000;
                    $note = 'Setengah hari / telat';
                } elseif ($val === '0') {
                    $status = 'tidak-hadir';
                    $deduction = 300000;
                    $note = 'Alpa';
                } elseif (strcasecmp($val, 'S') === 0 || stripos($val, 'sakit') !== false) {
                    $status = 'sakit';
                    $note = 'Sakit';
                } elseif (strcasecmp($val, 'I') === 0 || stripos($val, 'izin') !== false) {
                    $status = 'izin';
                    $note = 'Izin';
                } else {
                    $note = $val;
                    $status = 'hadir';
                }

                EmployeeAttendance::updateOrCreate(
                    [
                        'employee_id' => $employee->id,
                        'date' => $dateStr,
                    ],
                    [
                        'time_in' => $timeIn,
                        'time_out' => $timeOut,
                        'status' => $status,
                        'deduction' => $deduction,
                        'note' => $note,
                    ]
                );

                $result['attendance']++;
            }
        }
    }

    /**
     * Standard vertical tabular import (e.g. employee_id, name, date, time_in, time_out, status).
     */
    private function importVerticalSheet(Worksheet $sheet, array &$result, array &$affectedEmployeeIds): void
    {
        $rows = $sheet->toArray(null, true, false, true);

        if (count($rows) < 2) {
            $result['skipped']++;

            return;
        }

        $headerRowNumber = null;
        $headers = [];

        foreach (array_slice($rows, 0, 20, true) as $rowNumber => $row) {
            $candidate = [];

            foreach ($row as $column => $value) {
                $text = $this->normalizeHeader($value);
                if ($text !== '') {
                    $candidate[$column] = $text;
                }
            }

            $headerText = implode('|', array_values($candidate));

            $hasName = str_contains($headerText, 'nama') || str_contains($headerText, 'name') || str_contains($headerText, 'karyawan');
            $hasId = str_contains($headerText, 'id') || str_contains($headerText, 'nip') || str_contains($headerText, 'nik');

            if ($hasName || $hasId) {
                $headerRowNumber = $rowNumber;
                $headers = $candidate;
                break;
            }
        }

        if ($headerRowNumber === null) {
            $result['skipped']++;

            return;
        }

        $columnMap = [
            'date' => null,
            'employee_id' => null,
            'name' => null,
            'dept' => null,
            'time_in' => null,
            'time_out' => null,
            'status' => null,
            'deduction' => null,
            'note' => null,
        ];

        foreach ($headers as $column => $header) {
            if (in_array($header, ['tanggal', 'date', 'tgl'], true)) {
                $columnMap['date'] = $column;
            } elseif (in_array($header, ['id', 'employeeid', 'idkaryawan', 'kodekaryawan', 'nik', 'nip'], true)) {
                $columnMap['employee_id'] = $column;
            } elseif (in_array($header, ['nama', 'name', 'namakaryawan', 'karyawan'], true)) {
                $columnMap['name'] = $column;
            } elseif (in_array($header, ['dept', 'departemen', 'department', 'jabatan', 'divisi'], true)) {
                $columnMap['dept'] = $column;
            } elseif (in_array($header, ['masuk', 'jamasuk', 'checkin', 'timein', 'in'], true)) {
                $columnMap['time_in'] = $column;
            } elseif (in_array($header, ['pulang', 'jamkeluar', 'checkout', 'timeout', 'out'], true)) {
                $columnMap['time_out'] = $column;
            } elseif (in_array($header, ['status', 'kehadiran'], true)) {
                $columnMap['status'] = $column;
            } elseif (in_array($header, ['potongan', 'deduction', 'denda'], true)) {
                $columnMap['deduction'] = $column;
            } elseif (in_array($header, ['keterangan', 'catatan', 'note', 'remarks'], true)) {
                $columnMap['note'] = $column;
            }
        }

        if ($columnMap['name'] === null && $columnMap['employee_id'] === null) {
            $result['skipped']++;

            return;
        }

        foreach ($rows as $rowNumber => $row) {
            if ($rowNumber <= $headerRowNumber) {
                continue;
            }

            $name = $columnMap['name'] !== null ? trim((string) ($row[$columnMap['name']] ?? '')) : '';
            $employeeId = $columnMap['employee_id'] !== null ? trim((string) ($row[$columnMap['employee_id']] ?? '')) : null;

            if ($name === '' && empty($employeeId)) {
                continue;
            }

            if ($name === '' && ! empty($employeeId)) {
                $name = "Karyawan {$employeeId}";
            }

            $dept = $columnMap['dept'] !== null ? trim((string) ($row[$columnMap['dept']] ?? '')) : 'Umum';
            $timeIn = $columnMap['time_in'] !== null ? $this->parseTime($row[$columnMap['time_in']] ?? null) : null;
            $timeOut = $columnMap['time_out'] !== null ? $this->parseTime($row[$columnMap['time_out']] ?? null) : null;
            $rawDate = $columnMap['date'] !== null ? ($row[$columnMap['date']] ?? null) : null;
            $date = $this->parseDate($rawDate) ?? date('Y-m-d');
            $rawStatus = $columnMap['status'] !== null ? ($row[$columnMap['status']] ?? null) : null;
            $note = $columnMap['note'] !== null ? trim((string) ($row[$columnMap['note']] ?? '')) : null;

            $statusData = $this->detectStatusAndDeduction($timeIn, $timeOut, $note ?? $rawStatus, $date);

            if ($rawStatus !== null && trim((string) $rawStatus) !== '') {
                $statusData['status'] = $this->normalizeStatus($rawStatus);
            }

            if ($columnMap['deduction'] !== null && is_numeric($row[$columnMap['deduction']] ?? null)) {
                $statusData['deduction'] = (float) $row[$columnMap['deduction']];
            }

            $employee = $this->findOrCreateEmployee(
                $name,
                $employeeId,
                $dept !== '' ? $dept : 'Umum'
            );

            if ($employee->wasRecentlyCreated) {
                $result['employees']++;
            }

            $affectedEmployeeIds[] = $employee->id;

            EmployeeAttendance::updateOrCreate(
                [
                    'employee_id' => $employee->id,
                    'date' => $date,
                ],
                [
                    'time_in' => $timeIn,
                    'time_out' => $timeOut,
                    'status' => $statusData['status'],
                    'deduction' => $statusData['deduction'],
                    'note' => $note,
                ]
            );

            $result['attendance']++;
        }
    }

    /**
     * Synchronize Employee models with their latest attendance and aggregate deductions.
     *
     * @param  array<int>  $employeeIds
     */
    private function syncEmployees(array $employeeIds): void
    {
        if (empty($employeeIds)) {
            return;
        }

        foreach ($employeeIds as $id) {
            $employee = Employee::find($id);
            if (! $employee) {
                continue;
            }

            // Find attendance record with latest check_in or active presence
            $latest = EmployeeAttendance::where('employee_id', $id)
                ->where(function ($q) {
                    $q->whereNotNull('time_in')
                        ->orWhereIn('status', ['hadir', 'telat']);
                })
                ->orderBy('date', 'desc')
                ->first();

            if (! $latest) {
                $latest = EmployeeAttendance::where('employee_id', $id)
                    ->orderBy('date', 'desc')
                    ->first();
            }

            $totalDeduction = (float) EmployeeAttendance::where('employee_id', $id)->sum('deduction');

            $mappedStatus = 'belum-absen';
            if ($latest) {
                $mappedStatus = match ($latest->status) {
                    'hadir' => 'hadir',
                    'telat' => 'telat',
                    'izin', 'sakit', 'tidak-hadir' => 'tidak-hadir',
                    default => 'hadir',
                };
            }

            $employee->update([
                'status' => $mappedStatus,
                'check_in' => $latest?->time_in,
                'check_out' => $latest?->time_out,
                'deduction' => $totalDeduction,
            ]);
        }
    }

    /**
     * Finds or creates an Employee record.
     */
    public function findOrCreateEmployee(
        string $name,
        ?string $employeeId = null,
        ?string $dept = null
    ): Employee {
        $name = trim($name);
        $dept = ! empty($dept) ? trim($dept) : 'Umum';

        if ($employeeId !== null && trim($employeeId) !== '') {
            $cleanId = trim($employeeId);
            $existing = Employee::where('employee_id', $cleanId)->first();

            if ($existing) {
                if ($dept !== 'Umum' && ($existing->dept === null || $existing->dept === 'Umum')) {
                    $existing->update(['dept' => $dept]);
                }

                return $existing;
            }

            return Employee::create([
                'employee_id' => $cleanId,
                'name' => $name,
                'dept' => $dept,
                'status' => 'belum-absen',
                'deduction' => 0,
                'avatar' => $this->avatar($name),
            ]);
        }

        // Try lookup by exact name
        $byName = Employee::where('name', $name)->first();
        if ($byName) {
            if ($dept !== 'Umum' && ($byName->dept === null || $byName->dept === 'Umum')) {
                $byName->update(['dept' => $dept]);
            }

            return $byName;
        }

        $generatedId = 'EMP-'.strtoupper(
            substr(sha1(mb_strtolower($name)), 0, 8)
        );

        return Employee::firstOrCreate(
            ['employee_id' => $generatedId],
            [
                'name' => $name,
                'dept' => $dept,
                'status' => 'belum-absen',
                'deduction' => 0,
                'avatar' => $this->avatar($name),
            ]
        );
    }

    /**
     * Parses time values from strings, floats, decimals, etc.
     */
    public function parseTime(mixed $value): ?string
    {
        if ($value === null || $value === '') {
            return null;
        }

        if (is_numeric($value)) {
            $number = (float) $value;

            // Excel time fraction (0.0 to 0.999999)
            if ($number >= 0 && $number < 1) {
                $dateTime = ExcelDate::excelToDateTimeObject($number);

                return $dateTime->format('H:i:s');
            }

            // Integer hours (e.g. 17 -> 17:00:00)
            if ($number >= 1 && $number <= 24 && floor($number) == $number) {
                return sprintf('%02d:00:00', (int) $number);
            }
        }

        $text = trim((string) $value);
        $text = str_replace('.', ':', $text);

        // Check if string contains newline (take the first line)
        if (str_contains($text, "\n")) {
            $lines = array_filter(array_map('trim', explode("\n", $text)));
            if (! empty($lines)) {
                $text = reset($lines);
            }
        }

        $formats = [
            'H:i:s',
            'H:i',
            'g:i A',
            'h:i A',
            'g:i a',
            'h:i a',
            'H:i:s A',
            'H:i:s a',
        ];

        foreach ($formats as $format) {
            $date = DateTime::createFromFormat($format, $text);
            if ($date !== false) {
                return $date->format('H:i:s');
            }
        }

        // Try strtotime for natural time strings (e.g. "8:40 AM")
        if (preg_match('/\b\d{1,2}:\d{2}\b/', $text)) {
            $ts = strtotime($text);
            if ($ts !== false) {
                return date('H:i:s', $ts);
            }
        }

        return null;
    }

    /**
     * Parses date values.
     */
    public function parseDate(mixed $value): ?string
    {
        if ($value === null || $value === '') {
            return null;
        }

        if (is_numeric($value)) {
            try {
                return ExcelDate::excelToDateTimeObject($value)->format('Y-m-d');
            } catch (Throwable) {
                return null;
            }
        }

        $text = trim((string) $value);

        foreach (['Y-m-d', 'd/m/Y', 'd-m-Y', 'd.m.Y', 'd/m/y', 'd-m-y'] as $format) {
            $date = DateTime::createFromFormat($format, $text);
            if ($date !== false) {
                return $date->format('Y-m-d');
            }
        }

        return null;
    }

    private function extractDay(mixed $value): ?int
    {
        if ($value === null || $value === '') {
            return null;
        }

        $text = trim((string) $value);

        if (preg_match('/^\d{1,2}$/', $text)) {
            return (int) $text;
        }

        $date = $this->parseDate($value);

        return $date ? (int) date('d', strtotime($date)) : null;
    }

    /**
     * Detects status and computes late deduction.
     *
     * @return array{status: string, deduction: float, note: ?string}
     */
    public function detectStatusAndDeduction(?string $timeIn, ?string $timeOut, mixed $note = null, ?string $date = null): array
    {
        $noteText = $this->normalizeText($note);

        if (str_contains($noteText, 'izin') || str_contains($noteText, 'ijin') || str_contains($noteText, 'cuti')) {
            return [
                'status' => 'izin',
                'deduction' => 0,
                'note' => $note ? trim((string) $note) : 'Izin',
            ];
        }

        if (str_contains($noteText, 'sakit') || str_contains($noteText, 'mc')) {
            return [
                'status' => 'sakit',
                'deduction' => 0,
                'note' => $note ? trim((string) $note) : 'Sakit',
            ];
        }

        // If no check-in and no check-out
        if ($timeIn === null && $timeOut === null) {
            $isWeekend = false;
            if ($date) {
                $dayOfWeek = (int) date('w', strtotime($date));
                if ($dayOfWeek === 0) { // Sunday
                    $isWeekend = true;
                }
            }

            if ($isWeekend) {
                return [
                    'status' => 'tidak-hadir',
                    'deduction' => 0,
                    'note' => 'Libur Akhir Pekan',
                ];
            }

            return [
                'status' => 'tidak-hadir',
                'deduction' => 300000,
                'note' => $note ? trim((string) $note) : 'Alpa / Tidak Hadir',
            ];
        }

        // If timeIn is present
        if ($timeIn !== null) {
            $inMinutes = $this->timeToMinutes($timeIn);
            $workStartMinutes = 8 * 60 + 30; // 08:30
            $graceMinutes = 10;        // 08:40 tolerance (10 minutes after 08:30)

            // Standard morning shift
            if ($inMinutes <= $workStartMinutes + $graceMinutes) {
                return [
                    'status' => 'hadir',
                    'deduction' => 0,
                    'note' => $note ? trim((string) $note) : null,
                ];
            }

            // Morning late (after 08:15 and before 10:00)
            if ($inMinutes <= 10 * 60) {
                $lateMinutes = $inMinutes - $workStartMinutes;
                $deduction = $lateMinutes * 2000;

                return [
                    'status' => 'telat',
                    'deduction' => (float) $deduction,
                    'note' => $note ? trim((string) $note) : "Terlambat {$lateMinutes} menit",
                ];
            }

            // Afternoon or evening shift (e.g. 10:20 or 11:00)
            return [
                'status' => 'hadir',
                'deduction' => 0,
                'note' => $note ? trim((string) $note) : 'Shift Siang / Khusus',
            ];
        }

        // Only timeOut present
        return [
            'status' => 'hadir',
            'deduction' => 0,
            'note' => $note ? trim((string) $note) : 'Tidak ada jam masuk',
        ];
    }

    public function normalizeStatus(mixed $value): string
    {
        $text = $this->normalizeText($value);

        return match (true) {
            str_contains($text, 'telat'),
            str_contains($text, 'terlambat') => 'telat',

            str_contains($text, 'izin'),
            str_contains($text, 'ijin') => 'izin',

            str_contains($text, 'sakit') => 'sakit',

            str_contains($text, 'alpa'),
            str_contains($text, 'tidak hadir'),
            str_contains($text, 'absen') => 'tidak-hadir',

            default => 'hadir',
        };
    }

    private function normalizeHeader(mixed $value): string
    {
        $text = $this->normalizeText($value);

        return str_replace([' ', '-', '/', '.', '(', ')'], '', $text);
    }

    private function normalizeText(mixed $value): string
    {
        return mb_strtolower(trim((string) ($value ?? '')));
    }

    public function monthNumber(string $value): ?int
    {
        return match ($value) {
            'januari', 'january', 'jan' => 1,
            'februari', 'february', 'feb' => 2,
            'maret', 'march', 'mar', 'mrt' => 3,
            'april', 'apr' => 4,
            'mei', 'may' => 5,
            'juni', 'june', 'jun' => 6,
            'juli', 'july', 'jul' => 7,
            'agustus', 'august', 'aug', 'agt', 'agu' => 8,
            'september', 'sep' => 9,
            'oktober', 'october', 'oct', 'okt' => 10,
            'november', 'nov' => 11,
            'desember', 'december', 'dec', 'des' => 12,
            default => null,
        };
    }

    private function isNoiseName(string $name): bool
    {
        $text = $this->normalizeText($name);

        if ($this->monthNumber($text) !== null) {
            return true;
        }

        return in_array($text, [
            'nama',
            'tgl',
            'tanggal',
            'date',
            'in',
            'out',
            'no',
            'total',
            'keterangan',
            'ket',
            'bln',
            'bulan',
            'thn',
            'tahun',
        ], true);
    }

    private function avatar(string $name): string
    {
        $parts = preg_split('/\s+/', trim($name));

        if (! $parts) {
            return '?';
        }

        $first = strtoupper(substr($parts[0], 0, 1));
        $last = count($parts) > 1 ? strtoupper(substr($parts[count($parts) - 1], 0, 1)) : '';

        return $first.$last;
    }

    public function timeToMinutes(string $time): int
    {
        [$hour, $minute] = array_map('intval', explode(':', $time));

        return ($hour * 60) + $minute;
    }
}
