<?php

declare(strict_types=1);

namespace App\Helpers;

use App\Rules\TimeTableRule2023;

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

    public static function mapEmployersHours(): array
    {
        return [
            TimeTableRule2023::FREE_DAY_AFTER_NIGHT_TURN => TimeTableRule2023::MAX_NO_OF_DAILY_WORKING_HOURS,
            TimeTableRule2023::WORKING_ON_THE_MORNING => TimeTableRule2023::MAX_NO_OF_DAILY_WORKING_HOURS,
            TimeTableRule2023::WORKING_ON_THE_AFTERNOON => TimeTableRule2023::MAX_NO_OF_DAILY_WORKING_HOURS,
            TimeTableRule2023::WORKING_ON_THE_NIGHT => TimeTableRule2023::MAX_NO_OF_DAILY_WORKING_HOURS,
        ];
    }

    public static function mapColumnTimeTable(): array
    {
        return [
            1 => 'B',
            2 => 'C',
            3 => 'D',
            4 => 'E',
            5 => 'F',
            6 => 'G',
            7 => 'H',
            8 => 'I',
            9 => 'J',
            10 => 'K',
            11 => 'L',
            12 => 'M',
            13 => 'N',
            14 => 'O',
            15 => 'P',
            16 => 'Q',
            17 => 'R',
            18 => 'S',
            19 => 'T',
            20 => 'U',
            21 => 'V',
            22 => 'W',
            23 => 'X',
            24 => 'Y',
            25 => 'Z',
            26 => 'AA',
            27 => 'AB',
            28 => 'AC',
            29 => 'AD',
            30 => 'AE',
            31 => 'AF',
        ];
    }
}
