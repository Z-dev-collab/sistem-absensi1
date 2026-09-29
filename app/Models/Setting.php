<?php

namespace App\Models;

use Illuminate\Database\Eloquent\Factories\HasFactory;
use Illuminate\Database\Eloquent\Model;

class Setting extends Model
{
    use HasFactory;

    protected $fillable = [
        'start_time',
        'end_time',
        'grace_period',
        'rate_per_minute',
        'radius_meters',
    ];
}
