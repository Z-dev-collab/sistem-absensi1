<?php

use Illuminate\Database\Migrations\Migration;
use Illuminate\Database\Schema\Blueprint;
use Illuminate\Support\Facades\Schema;

return new class extends Migration
{
    public function up(): void
    {
        Schema::create('employees', function (Blueprint $table) {
            $table->id();
            $table->string('employee_id')->unique();
            $table->string('name');
            $table->string('dept')->nullable();

            $table->enum('status', [
                'hadir',
                'telat',
                'tidak-hadir',
                'belum-absen',
            ])->default('belum-absen');

            $table->time('check_in')->nullable();
            $table->time('check_out')->nullable();
            $table->decimal('deduction', 15, 2)->default(0);
            $table->string('avatar', 10)->nullable();

            $table->timestamps();
        });
    }

    public function down(): void
    {
        Schema::dropIfExists('employees');
    }
};
