<?php

namespace Database\Seeders;

use Illuminate\Database\Seeder;

class DatabaseSeeder extends Seeder
{
    /**
     * Seed the application's database.
     */
    public function run(): void
    {
        // Memanggil SettingSeeder yang sudah kamu buat sebelumnya
        $this->call([
            SettingSeeder::class,
        ]);
    }
}
