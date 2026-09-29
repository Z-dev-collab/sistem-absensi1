<?php

namespace App\Models;

use Illuminate\Database\Eloquent\Model;
use Illuminate\Database\Eloquent\Relations\HasMany;

class Employee extends Model
{
    protected $fillable = [
        'employee_id',
        'name',
        'dept',
        'status',
        'check_in',
        'check_out',
        'deduction',
        'avatar',
    ];

    protected $casts = [
        'deduction' => 'decimal:2',
    ];

    public function attendances(): HasMany
    {
        return $this->hasMany(EmployeeAttendance::class);
    }
}
