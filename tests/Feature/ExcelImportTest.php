<?php

namespace Tests\Feature;

use App\Models\Employee;
use Illuminate\Http\UploadedFile;
use Tests\TestCase;

class ExcelImportTest extends TestCase
{
    public function test_can_fetch_employees_list(): void
    {
        $response = $this->getJson(route('employees.index'));

        $response->assertOk();
        $response->assertJsonIsArray();
    }

    public function test_can_import_maintenance_excel(): void
    {
        $samplePath = base_path('storage/app/sample-excel/ABSENSI TEAM MAINTENENCE(1).xlsx');

        if (! file_exists($samplePath)) {
            $this->markTestSkipped('Sample maintenance excel not available.');
        }

        $file = new UploadedFile(
            $samplePath,
            'ABSENSI TEAM MAINTENENCE(1).xlsx',
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            null,
            true
        );

        $response = $this->postJson(route('employees.import'), [
            'file' => $file,
        ]);

        $response->assertOk()
            ->assertJsonStructure([
                'message',
                'count',
                'attendance_count',
            ]);

        $this->assertDatabaseHas('employees', [
            'name' => 'TRI SUPARTMAN',
        ]);
    }

    public function test_can_import_fingerprint_attendance_excel(): void
    {
        $samplePath = base_path('storage/app/sample-excel/Absen Agust 1-30 Agust 26.xls');

        if (! file_exists($samplePath)) {
            $this->markTestSkipped('Sample fingerprint excel not available.');
        }

        $file = new UploadedFile(
            $samplePath,
            'Absen Agust 1-30 Agust 26.xls',
            'application/vnd.ms-excel',
            null,
            true
        );

        $response = $this->postJson(route('employees.import'), [
            'file' => $file,
        ]);

        $response->assertOk()
            ->assertJsonStructure([
                'message',
                'count',
                'attendance_count',
            ]);

        $this->assertDatabaseHas('employees', [
            'name' => 'adam husein',
        ]);
    }

    public function test_import_rejects_invalid_file(): void
    {
        $file = UploadedFile::fake()->create('document.pdf', 100, 'application/pdf');

        $response = $this->postJson(route('employees.import'), [
            'file' => $file,
        ]);

        $response->assertStatus(422)
            ->assertJsonValidationErrors(['file']);
    }

    public function test_can_delete_single_employee(): void
    {
        $emp = Employee::create([
            'employee_id' => 'EMP-TEST-999',
            'name' => 'Testing Delete One',
            'dept' => 'QA',
            'status' => 'belum-absen',
            'deduction' => 0,
        ]);

        $response = $this->deleteJson(route('employees.destroy', ['employee' => $emp->id]));

        $response->assertOk()
            ->assertJsonFragment(['message' => 'Data karyawan Testing Delete One berhasil dihapus.']);

        $this->assertDatabaseMissing('employees', ['id' => $emp->id]);
    }

    public function test_can_bulk_delete_employees(): void
    {
        $emp1 = Employee::create([
            'employee_id' => 'EMP-BULK-1',
            'name' => 'Bulk Emp 1',
            'dept' => 'QA',
            'status' => 'belum-absen',
            'deduction' => 0,
        ]);

        $emp2 = Employee::create([
            'employee_id' => 'EMP-BULK-2',
            'name' => 'Bulk Emp 2',
            'dept' => 'QA',
            'status' => 'belum-absen',
            'deduction' => 0,
        ]);

        $response = $this->postJson(route('employees.bulk-destroy'), [
            'ids' => [$emp1->id, $emp2->id],
        ]);

        $response->assertOk()
            ->assertJsonFragment(['deleted_count' => 2]);

        $this->assertDatabaseMissing('employees', ['id' => $emp1->id]);
        $this->assertDatabaseMissing('employees', ['id' => $emp2->id]);
    }
}
