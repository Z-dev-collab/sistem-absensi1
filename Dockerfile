FROM php:8.4-cli

RUN apt-get update && apt-get install -y \
    git \
    unzip \
    libpq-dev \
    libzip-dev \
    nodejs \
    npm \
    && docker-php-ext-install pdo pdo_pgsql \
    && rm -rf /var/lib/apt/lists/*

COPY --from=composer:2 /usr/bin/composer /usr/bin/composer

WORKDIR /var/www/html

COPY . .

RUN composer install --no-dev --optimize-autoloader --no-interaction

RUN npm install --no-audit --no-fund

RUN npm run build

RUN php artisan config:clear

ENV APP_ENV=production
ENV APP_DEBUG=false

EXPOSE 10000

CMD php artisan serve --host=0.0.0.0 --port=${PORT:-10000}