<?php

declare(strict_types=1);

namespace App\Rules;

class TimeTableRule2023
{
    public const MAX_NO_OF_NIGHTS = 3;
    public const MIN_NO_OF_MORNINGS = 3;
    public const MIN_NO_OF_AFTERNOONS = 3;
    public const MAX_NO_OF_PERSONS_FROM_MONDAY_TO_THURSDAY_NIGHTS = 2;
    public const MAX_NO_OF_PERSONS_FROM_ONE_TURN_WEEKEND = 1;
    public const MIN_NO_OF_WEEKEND_WORKING_HOURS = 12;
    public const MAX_NO_OF_WEEKEND_WORKING_HOURS = 24;
    public const MIN_NO_OF_PERSONS_FOR_MORNING_TURN = 6;
    public const MAX_NO_OF_PERSONS_FOR_MORNING_TURN = 9;
    public const MIN_NO_OF_PERSONS_FOR_AFTERNOON_TURN = 2;
    public const MAX_NO_OF_PERSONS_FOR_AFTERNOON_TURN = 3;
    public const MAX_NO_OF_WEEKLY_WORKING_HOURS = 30;
    public const MAX_NO_OF_DAILY_WORKING_HOURS = 6;

    public const WORKING_LAST_MONTH_NIGHT = [
        'DA' => true,
        'NU' => false,
    ];

    public static array $publicHolidays = [
        '2023-01-01',
        '2023-01-02',
        '2023-01-24',
        '2023-04-14',
        '2023-04-16',
        '2023-04-17',
        '2023-05-01',
        '2023-06-01',
        '2023-06-04',
        '2023-06-05',
        '2023-08-15',
        '2023-11-30',
        '2023-12-01',
        '2023-12-25',
        '2023-12-26',
    ];

    public const FREE_DAY = 'L';
    public const FREE_NATIONAL_DAY = 'LN';
    public const HOLIDAY = 'CO';
    public const WORKING_ON_THE_MORNING = 'D';
    public const WORKING_ON_THE_AFTERNOON = 'DM';
    public const WORKING_ON_THE_MORNING_OR_ON_THE_AFTERNOON = 'D|DM';
    public const WORKING_ON_THE_NIGHT = 'N';
    public const FREE_DAY_AFTER_NIGHT_TURN = '-';
    public const FREE_ON_WEEKEND_DAY = 'LW';
    public const PRIORITY_EMPLOYER = 'Pintican';

    public const WORKING_ON_WEEKENDS = [
        'DA' => true,
        'NU' => false,
    ];
}
