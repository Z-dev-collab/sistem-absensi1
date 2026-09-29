<?php

namespace App\Imports;

use App\Models\Employee;
use Illuminate\Database\Eloquent\Model;
use Maatwebsite\Excel\Concerns\ToModel;
use Maatwebsite\Excel\Concerns\WithHeadingRow;

class EmployeeImport implements ToModel, WithHeadingRow
{
    public function model(array $row): Model|array|null
    {
        // Bisa membaca format:
        // employee_id / id
        $employeeId = $row['employee_id']
            ?? $row['id']
            ?? null;

        // Bisa membaca format:
        // name / nama_karyawan
        $name = $row['name']
            ?? $row['nama_karyawan']
            ?? null;

        // Bisa membaca format:
        // dept / departemen
        $dept = $row['dept']
            ?? $row['departemen']
            ?? null;

        if (empty($employeeId) || empty($name)) {
            return null;
        }

        // Status
        $status = $row['status']
            ?? 'belum-absen';

        // Jam masuk
        $checkIn = $row['check_in']
            ?? $row['jam_masuk']
            ?? null;

        // Jam pulang
        $checkOut = $row['check_out']
            ?? $row['jam_pulang']
            ?? null;

        // Potongan
        $deduction = $row['deduction']
            ?? $row['total_potongan']
            ?? 0;

        // Avatar
        $avatar = $row['avatar']
            ?? null;

        return Employee::updateOrCreate(
            [
                'employee_id' => trim((string) $employeeId),
            ],
            [
                'name' => trim((string) $name),

                'dept' => ! empty($dept)
                    ? trim((string) $dept)
                    : null,

                'status' => trim((string) $status),

                'check_in' => ! empty($checkIn)
                    ? $checkIn
                    : null,

                'check_out' => ! empty($checkOut)
                    ? $checkOut
                    : null,

                'deduction' => ! empty($deduction)
                    ? (float) $deduction
                    : 0,

                'avatar' => ! empty($avatar)
                    ? trim((string) $avatar)
                    : null,
            ]
        );
    }
}
