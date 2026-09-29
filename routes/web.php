<?php

use App\Http\Controllers\EmployeeController;
use Illuminate\Support\Facades\Route;

Route::get('/', function () {
    return view('app');
});

Route::get('/employees', [EmployeeController::class, 'index'])
    ->name('employees.index');

Route::post('/employees/import', [EmployeeController::class, 'import'])
    ->name('employees.import');

Route::delete('/employees/all', [EmployeeController::class, 'destroyAll'])
    ->name('employees.destroy-all');

Route::post('/employees/bulk-delete', [EmployeeController::class, 'bulkDestroy'])
    ->name('employees.bulk-destroy');

Route::delete('/employees/{employee}', [EmployeeController::class, 'destroy'])
    ->name('employees.destroy');
