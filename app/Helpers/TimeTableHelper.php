<?php

declare(strict_types=1);

namespace App\Helpers;

class TimeTableHelper
{
    public static function removeEmptyValuesFromDateInterval(string $dateInterval): array
    {
        return array_filter(
            explode(
                ',',
                preg_replace('/\s+/', '', $dateInterval) ?? ''
            ),
            fn(mixed $value) => !in_array($value, ['', null])
        );
    }
}
