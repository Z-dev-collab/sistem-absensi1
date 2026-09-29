<?php

namespace App\Http\Controllers;

use App\Models\Employee;
use App\Services\AttendanceExcelImportService;
use Illuminate\Http\JsonResponse;
use Illuminate\Http\RedirectResponse;
use Illuminate\Http\Request;
use Throwable;

class EmployeeController extends Controller
{
    public function __construct(
        public AttendanceExcelImportService $importService
    ) {}

    /**
     * Get list of employees with attendance details.
     */
    public function index(Request $request): JsonResponse
    {
        $date = $request->query('date');

        $query = Employee::orderBy('name');

        if ($date) {
            $employees = $query->with(['attendances' => function ($q) use ($date) {
                $q->where('date', $date);
            }])->get()->map(function (Employee $emp) {
                $att = $emp->attendances->first();

                return [
                    'id' => $emp->id,
                    'employee_id' => $emp->employee_id,
                    'name' => $emp->name,
                    'dept' => $emp->dept ?? 'Umum',
                    'status' => $att?->status ?? 'belum-absen',
                    'check_in' => $att?->time_in ? substr($att->time_in, 0, 5) : null,
                    'check_out' => $att?->time_out ? substr($att->time_out, 0, 5) : null,
                    'deduction' => (float) ($att?->deduction ?? 0),
                    'avatar' => $emp->avatar,
                ];
            });

            return response()->json($employees);
        }

        return response()->json($query->get());
    }

    /**
     * Import attendance Excel file (.xlsx, .xls, .csv).
     */
    public function import(Request $request): JsonResponse|RedirectResponse
    {
        $validated = $request->validate([
            'file' => [
                'required',
                'file',
                'extensions:xlsx,xls,csv',
                'max:15360',
            ],
        ]);

        try {
            $file = $validated['file'];
            $realPath = $file->getRealPath();

            $result = $this->importService->import($realPath);

            $message = sprintf(
                'Data Excel berhasil diimport (%d karyawan, %d data absensi).',
                $result['total_employees'],
                $result['attendance']
            );

            if ($request->expectsJson()) {
                return response()->json([
                    'message' => $message,
                    'count' => $result['total_employees'],
                    'new_employees' => $result['employees'],
                    'attendance_count' => $result['attendance'],
                    'sheets' => $result['sheets'],
                    'errors' => $result['errors'],
                ]);
            }

            return back()->with('success', $message);
        } catch (Throwable $e) {
            if ($request->expectsJson()) {
                return response()->json([
                    'message' => 'Gagal mengimport file Excel: '.$e->getMessage(),
                ], 422);
            }

            return back()->withErrors(['file' => 'Gagal mengimport file Excel: '.$e->getMessage()]);
        }
    }

    /**
     * Delete a single employee (one for one).
     */
    public function destroy(string $id): JsonResponse
    {
        $employee = Employee::where('id', $id)
            ->orWhere('employee_id', $id)
            ->first();

        if (! $employee) {
            return response()->json(['message' => 'Karyawan tidak ditemukan.'], 404);
        }

        $name = $employee->name;
        $employee->delete();

        return response()->json([
            'message' => "Data karyawan {$name} berhasil dihapus.",
            'count' => Employee::count(),
        ]);
    }

    /**
     * Delete selected employees (bulk delete).
     */
    public function bulkDestroy(Request $request): JsonResponse
    {
        $validated = $request->validate([
            'ids' => ['required', 'array'],
            'ids.*' => ['required'],
        ]);

        $ids = $validated['ids'];
        $count = Employee::whereIn('id', $ids)
            ->orWhereIn('employee_id', $ids)
            ->delete();

        return response()->json([
            'message' => "{$count} data karyawan berhasil dihapus.",
            'deleted_count' => $count,
            'count' => Employee::count(),
        ]);
    }

    /**
     * Delete all employee data.
     */
    public function destroyAll(): JsonResponse
    {
        $count = Employee::count();
        Employee::query()->delete();

        return response()->json([
            'message' => "Semua ({$count}) data karyawan berhasil dihapus.",
            'deleted_count' => $count,
            'count' => 0,
        ]);
    }
}
