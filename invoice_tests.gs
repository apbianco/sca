function testNormalizeName() {
  var testCases = [
    {
      description: "Input with leading/trailing spaces",
      input: "  Test Name  ",
      expected: "Test-Name",
      to_upper_case: false
    },
    {
      description: "Input with diacritics",
      input: "Élève Test",
      expected: "Eleve-Test",
      to_upper_case: false
    },
    {
      description: "Input with special characters",
      input: "Test/Name_One.Two",
      expected: "Test-Name-One-Two",
      to_upper_case: false
    },
    {
      description: "Input for uppercase conversion",
      input: "test name",
      expected: "TEST-NAME",
      to_upper_case: true
    },
    {
      description: "Input with mixed special characters and spaces",
      input: "  Test/Name_One.Two  ",
      expected: "Test-Name-One-Two",
      to_upper_case: false
    },
    {
      description: "Input with diacritics and uppercase conversion",
      input: "Élève Test",
      expected: "ELEVE-TEST",
      to_upper_case: true
    },
    {
      description: "Input with numbers (should be removed)",
      input: "Test123Name456",
      expected: "TestName",
      to_upper_case: false
    },
    {
      description: "Input with multiple hyphens (should be reduced to one)",
      input: "Test---Name",
      expected: "Test-Name",
      to_upper_case: false
    },
    {
      description: "Input with trailing hyphen (should be removed)",
      input: "Test-Name-",
      expected: "Test-Name",
      to_upper_case: false
    }
  ];

  var failures = 0;

  for (var i = 0; i < testCases.length; i++) {
    var tc = testCases[i];
    var actual = normalizeName(tc.input, tc.to_upper_case);
    if (actual !== tc.expected) {
      failures++;
      Logger.log("--------------------------------------------------");
      Logger.log("FAIL: testNormalizeName - " + tc.description);
      Logger.log("Input: '" + tc.input + "' (to_upper_case: " + tc.to_upper_case + ")");
      Logger.log("Expected: '" + tc.expected + "'");
      Logger.log("Got: '" + actual + "'");
    }
  }
  
  if (failures > 0) {
    Logger.log("--------------------------------------------------");
    Logger.log("testNormalizeName: " + failures + " test(s) failed.");
  }
  // No summary log if all passed, just the "Finished" log.
  Logger.log("Finished testNormalizeName().");
  return failures === 0;
}

function testPlural() {
  var testCases = [
    {
      description: "Singular input",
      number: 1,
      message: "item",
      expected: "item"
    },
    {
      description: "Plural input",
      number: 2,
      message: "item",
      expected: "items"
    },
    {
      description: "Zero input",
      number: 0,
      message: "item",
      expected: "item"
    },
    {
      description: "Multi-word input 1",
      number: 3,
      message: "apple tree",
      expected: "apples trees"
    },
    {
      description: "Multi-word input 2",
      number: 3,
      message: "apple tree top",
      expected: "apples trees tops"
    },
    {
      description: "Input with leading/trailing spaces, singular",
      number: 1,
      message: " apple tree ",
      expected: " apple tree "
    },
    {
      description: "Input with leading/trailing spaces, plural",
      number: 2,
      message: " apple tree ",
      expected: " apples trees "
    },
    {
      description: "Input with only one space, singular",
      number: 1,
      message: " ",
      expected: " "
    },
    {
      description: "Input with only spaces, singular",
      number: 1,
      message: "     ",
      expected: "     "
    },    
    {
      description: "Input with only one space, plural",
      number: 2,
      message: " ",
      expected: " "
    },
    {
      description: "Input with only spaces, plural",
      number: 2,
      message: "   ",
      expected: "   "
    },
    {
      description: "Empty input, singular",
      number: 1,
      message: "",
      expected: ""
    },
    {
      description: "Empty input, plural",
      number: 2,
      message: "",
      expected: ""
    },
  ];

  var failures = 0;

  for (var i = 0; i < testCases.length; i++) {
    var tc = testCases[i];
    var actual = Plural(tc.number, tc.message);
    if (actual !== tc.expected) {
      failures++;
      Logger.log("--------------------------------------------------");
      Logger.log("FAIL: testPlural - " + tc.description);
      Logger.log("Input number: " + tc.number + ", Input message: '" + tc.message + "'");
      Logger.log("Expected: '" + tc.expected + "'");
      Logger.log("Got: '" + actual + "'");
    }
  }

  if (failures > 0) {
    Logger.log("--------------------------------------------------");
    Logger.log("testPlural: " + failures + " test(s) failed.");
  }
  // No summary log if all passed, just the "Finished" log.
  Logger.log("Finished testPlural().");
  return failures === 0;
}

function testGetDoBYear() {
  var testCases = [
    {
      description: "Date string (MM/DD/YYYY)",
      input: "12/31/1995",
      input_type: "string",
      expected: 1995
    },
    {
      description: "JavaScript Date object (Month is 0-indexed)",
      input: new Date(2005, 0, 1), // Month 0 is January
      input_type: "object",
      expected: 2005
    },
    {
      description: "Date string (YYYY-MM-DD)",
      input: "1987-05-15",
      input_type: "string",
      expected: 1987
    },
    {
      description: "Date string (M/D/YYYY without leading zeros)",
      input: "1/5/2003",
      input_type: "string",
      expected: 2003
    },
    {
      description: "JavaScript Date object (leap year)",
      input: new Date(2000, 1, 29), // Feb 29, 2000
      input_type: "object",
      expected: 2000
    }
  ];

  var failures = 0;

  for (var i = 0; i < testCases.length; i++) {
    var tc = testCases[i];
    var actual = getDoBYear(tc.input);
    if (actual !== tc.expected) {
      failures++;
      Logger.log("--------------------------------------------------");
      Logger.log("FAIL: testGetDoBYear - " + tc.description);
      var inputString = "";
      if (tc.input_type === "string") {
        inputString = "\"" + tc.input + "\"";
      } else {
        // For Date objects, construct a string representation
        inputString = "new Date(" + tc.input.getFullYear() + ", " + tc.input.getMonth() + ", " + tc.input.getDate() + ")";
      }
      Logger.log("Input: " + inputString);
      Logger.log("Expected: " + tc.expected);
      Logger.log("Got: " + actual);
    }
  }

  if (failures > 0) {
    Logger.log("--------------------------------------------------");
    Logger.log("testGetDoBYear: " + failures + " test(s) failed.");
  }
  // No summary log if all passed, just the "Finished" log.
  Logger.log("Finished testGetDoBYear().");
  return failures === 0;
}

function testIsLicenseDefined() {
  var testCases = [
    {
      description: "Empty string",
      input: "",
      expected: false
    },
    {
      description: "String 'Aucune'",
      input: "Aucune",
      expected: false
    },
    {
      description: "String that is not a license but not empty or 'Aucune'",
      input: "InvalidLicense",
      expected: false
    },
    {
      description: "Typical valid license string",
      input: "CN Jeune (Loisir)",
      expected: true
    },
    {
      description: "License string",
      input: "CN Adulte (Loisir)",
      expected: true
    },
    {
      description: "License string",
      input: "CN Dirigeant",
      expected: true
    }
  ];

  var failures = 0;

  for (var i = 0; i < testCases.length; i++) {
    var tc = testCases[i];
    // Assuming isLicenseDefined is globally accessible
    var actual = isLicenseDefined(tc.input);
    if (actual !== tc.expected) {
      failures++;
      Logger.log("--------------------------------------------------");
      Logger.log("FAIL: testIsLicenseDefined - " + tc.description);
      Logger.log("Input: '" + tc.input + "'");
      Logger.log("Expected: " + tc.expected);
      Logger.log("Got: '" + actual + "'");
    }
  }

  if (failures > 0) {
    Logger.log("--------------------------------------------------");
    Logger.log("testIsLicenseDefined: " + failures + " test(s) failed.");
  }
  // No summary log if all passed, just the "Finished" log.
  Logger.log("Finished testIsLicenseDefined().");
  return failures === 0;
}

function testAgeFromDoB() {
  var testCases = [];
  var today = new Date();
  var currentYear = today.getFullYear();
  var currentMonth = today.getMonth();
  var currentDate = today.getDate();

  // Test Case 1: Exactly 18 years ago
  testCases.push({
    description: "Exactly 18 years ago",
    input: new Date(currentYear - 18, currentMonth, currentDate),
    expected: 18
  });

  // Test Case 2: Less than 18 years ago (e.g., 10 years ago)
  testCases.push({
    description: "10 years ago",
    input: new Date(currentYear - 10, currentMonth, currentDate),
    expected: 10
  });

  // Test Case 3: More than 18 years ago (e.g., 30 years ago)
  testCases.push({
    description: "30 years ago",
    input: new Date(currentYear - 30, currentMonth, currentDate),
    expected: 30
  });

  // Test Case 4: Leap year birthday (Feb 29, 2000)
  var leapDob = new Date(2000, 1, 29); // Month is 0-indexed, so 1 is February
  var expectedAgeLeap;
  // Calculate expected age for leap year case
  var ageDateLeap = new Date(Date.now() - leapDob.getTime());
  expectedAgeLeap = Math.abs(ageDateLeap.getUTCFullYear() - 1970);
  testCases.push({
    description: "Leap year birthday (Feb 29, 2000)",
    input: leapDob,
    expected: expectedAgeLeap
  });

  // Test Case 5: Today's date
  testCases.push({
    description: "Today's date",
    input: new Date(), // A new Date object for today
    expected: 0
  });

  // Test Case 6: Future date (e.g., 1 year in future)
  var futureDate = new Date(currentYear + 1, currentMonth, currentDate);
  var month_diff_future = Date.now() - futureDate.getTime();
  var age_dt_future = new Date(month_diff_future);
  var year_future = age_dt_future.getUTCFullYear();
  testCases.push({
    description: "1 year in the future",
    input: futureDate,
    expected: -1
  });

  var failures = 0;

  for (var i = 0; i < testCases.length; i++) {
    var tc = testCases[i];
    var actual = ageFromDoB(tc.input); // Assuming ageFromDoB is globally accessible
    if (actual !== tc.expected) {
      // Add a small tolerance for "Today's date" due to execution time lag
      if (tc.description === "Today's date" && actual === 0 && tc.expected === 0) {
        // This is fine, often dob.getTime() can be slightly different from Date.now()
      } else if (tc.description.includes("years ago") && tc.input.getMonth() === currentMonth && tc.input.getDate() === currentDate) {
        // For "exactly N years ago" cases, if the test runs very close to midnight,
        // Date.now() might shift to the next day while tc.input is fixed.
        // This could cause a 1-year difference if not handled.
        // A more robust way would be to check if actual is tc.expected or tc.expected - 1
        // but for now, we'll assume tests don't run exactly at midnight causing this specific discrepancy.
        // The current logic of ageFromDoB is based on millisecond difference, so it should be fairly precise.
      }
      // For the leap year, the expected age is calculated dynamically.
      // For future dates, the expected age is also calculated dynamically.
      // Allow a small tolerance for floating point comparisons if ages were not integers.
      // However, ageFromDoB returns integers, so direct comparison should be fine.

      // Check if the difference is due to the "day boundary" issue for exact year differences
      // If the calculated age is one less than expected, and the birth month/day is today,
      // it might be due to the time component making the person not "fully" N years old yet.
      // The ageFromDoB function calculates age purely based on (Current Time - DoB Time) converted to years.
      // Example: If today is Nov 17, 2023, 10:00 AM.
      // DoB: Nov 17, 2005, 11:00 AM. Age is 17, not 18 yet. ageFromDoB will give 17.
      // DoB: Nov 17, 2005, 09:00 AM. Age is 18. ageFromDoB will give 18.
      // The test cases using currentMonth and currentDate for "exactly N years ago" assume the time component
      // results in the person being fully N years old.

      // A simple log for now. More complex tolerance logic can be added if needed.
      if (actual !== tc.expected) {
        failures++;
        Logger.log("--------------------------------------------------");
        Logger.log("FAIL: testAgeFromDoB - " + tc.description);
        Logger.log("Input Date: " + tc.input.toString());
        Logger.log("Expected: " + tc.expected);
        Logger.log("Got: " + actual);
        // Log for debugging future date calculation
        if (tc.description.includes("future")) {
            Logger.log("DEBUG Future Date: month_diff_future=" + month_diff_future +
                       ", age_dt_future=" + age_dt_future.toUTCString() +
                       ", year_future=" + year_future);
        }
      }
    }
  }

  if (failures > 0) {
    Logger.log("--------------------------------------------------");
    Logger.log("testAgeFromDoB: " + failures + " test(s) failed.");
  }
  Logger.log("Finished testAgeFromDoB().");
  return failures === 0;
}

function testFormatPhoneNumberString() {
  var testCases = [
    { description: "Already formatted", input: "01 23 45 67 89", expected: "01 23 45 67 89" },
    { description: "Numbers with hyphens", input: "01-23-45-67-89", expected: "01 23 45 67 89" },
    { description: "Numbers with mixed spaces and hyphens", input: "01 23-45 67-89", expected: "01 23 45 67 89" },
    { description: "Numbers with no spaces or hyphens", input: "0123456789", expected: "01 23 45 67 89" },
    { description: "Numbers with leading/trailing internal spaces (handled by replace all spaces)", input: " 0123456789 ", expected: "01 23 45 67 89" },
    { description: "Already formatted with extraneous internal spaces", input: "01  23   45--67  89", expected: "01 23 45 67 89" },
    { description: "Empty string", input: "", expected: "" },
    { description: "Null input", input: null, expected: "" },
    { description: "Short number", input: "01234", expected: "01 23 4" },
    { description: "Odd length number", input: "0123456", expected: "01 23 45 6" },
    { description: "Number with internal single digits (robustness)", input: "01 2 34 5 67", expected: "01 23 45 67" }
  ];

  var failures = 0;

  for (var i = 0; i < testCases.length; i++) {
    var tc = testCases[i];
    var actual = formatPhoneNumberString(tc.input); // Assuming formatPhoneNumberString is globally accessible
    if (actual !== tc.expected) {
      failures++;
      Logger.log("--------------------------------------------------");
      Logger.log("FAIL: testFormatPhoneNumberString - " + tc.description);
      Logger.log("Input: '" + tc.input + "' (type: " + typeof tc.input + ")");
      Logger.log("Expected: '" + tc.expected + "'");
      Logger.log("Got: '" + actual + "'");
    }
  }

  if (failures > 0) {
    Logger.log("--------------------------------------------------");
    Logger.log("testFormatPhoneNumberString: " + failures + " test(s) failed.");
  }
  Logger.log("Finished testFormatPhoneNumberString().");
  return failures === 0;
}

/////////////////////////////////////////////////////////////////////////////////////////
// config.gs tests
/////////////////////////////////////////////////////////////////////////////////////////

function testSeasonConfiguration() {
  var passed = true;
  var currentYear = new Date().getFullYear();
  var expectedSeason = currentYear + "/" + (currentYear + 1);
  var actualSeason = season; // From config.gs

  if (actualSeason !== expectedSeason) {
    passed = false;
    Logger.log("--------------------------------------------------");
    Logger.log("FAIL: testSeasonConfiguration - Season mismatch");
    Logger.log("  Expected season for current year (" + currentYear + "): '" + expectedSeason + "'");
    Logger.log("  Actual season configured in config.gs: '" + actualSeason + "'");
  }
  Logger.log("Finished testSeasonConfiguration().");
  return passed;
}

function testLicensesConfiguration() {
  Logger.log("Starting testLicensesConfiguration()...");
  var allTestsPassed = true;
  var simulatedFutureYear = new Date().getFullYear() + 1;
  var maxJeuneAge = 18; // As defined in previous logic

  // licenses_configuration_map from config.gs
  for (var licenseName in licenses_configuration_map) {
    if (licenses_configuration_map.hasOwnProperty(licenseName)) {
      var configuredBirthYear = licenses_configuration_map[licenseName];

      if (licenseName.indexOf('Jeune') !== -1) {
        var ageInSimulatedFutureYear = simulatedFutureYear - configuredBirthYear;
        if (ageInSimulatedFutureYear > maxJeuneAge) {
          allTestsPassed = false;
          Logger.log("--------------------------------------------------");
          Logger.log("FAIL: testLicensesConfiguration - Outdated 'Jeune' category");
          Logger.log("  License: '" + licenseName + "'");
          Logger.log("  Configured Birth Year: " + configuredBirthYear);
          Logger.log("  Simulated Future Year: " + simulatedFutureYear);
          Logger.log("  Calculated Age in Future Year: " + ageInSimulatedFutureYear + ", exceeds max " + maxJeuneAge);
        }
      } else if (licenseName.indexOf('Adulte') !== -1 || licenseName.indexOf('Dirigeant') !== -1) {
        if (configuredBirthYear >= simulatedFutureYear) {
          allTestsPassed = false;
          Logger.log("--------------------------------------------------");
          Logger.log("FAIL: testLicensesConfiguration - Illogical birth year for 'Adulte'/'Dirigeant'");
          Logger.log("  License: '" + licenseName + "'");
          Logger.log("  Configured Birth Year: " + configuredBirthYear);
          Logger.log("  Simulated Future Year: " + simulatedFutureYear);
          Logger.log("  Birth year must be less than simulated future year.");
        }
      }
    }
  }

  // New checks for row and column validity
  for (var categoryName in skipass_configuration_map) {
    if (skipass_configuration_map.hasOwnProperty(categoryName)) {
      var configValue = skipass_configuration_map[categoryName];

      if (!configValue || typeof configValue.length === 'undefined') {
        allTestsPassed = false;
        Logger.log("--------------------------------------------------");
        Logger.log("FAIL: testSkipassConfiguration - Invalid or undefined configValue for '" + categoryName + "'");
        continue; // Skip further checks for this malformed entry
      }

      if (configValue.length !== 4) {
        allTestsPassed = false;
        Logger.log("--------------------------------------------------");
        Logger.log("FAIL: testSkipassConfiguration - Incorrect array length for '" + categoryName + "'");
        Logger.log("  Expected 4 elements (Dob/Y1, Dob/Y2, Row, Col). Got: " + configValue.length + " -> [" + configValue.join(", ") + "]");
      } else {
        // Only check row and col if length is potentially correct
        var row = configValue[2];
        var col = configValue[3];
        if (typeof row !== 'number' || row <= 0) {
          allTestsPassed = false;
          Logger.log("--------------------------------------------------");
          Logger.log("FAIL: testSkipassConfiguration - Invalid row for '" + categoryName + "'");
          Logger.log("  Row value: " + row + " (type: " + typeof row + "). Expected a positive number. Array: [" + configValue.join(", ") + "]");
        }
        if (typeof col !== 'number' || col <= 0) {
          allTestsPassed = false;
          Logger.log("--------------------------------------------------");
          Logger.log("FAIL: testSkipassConfiguration - Invalid column for '" + categoryName + "'");
          Logger.log("  Column value: " + col + " (type: " + typeof col + "). Expected a positive number. Array: [" + configValue.join(", ") + "]");
        }
      }
    }
  }
  // End of new checks
  Logger.log("Finished testLicensesConfiguration().");
  return allTestsPassed;
}

function testCompSubscriptionConfiguration() {
  var allTestsPassed = true;
  var simulatedFutureYear = new Date().getFullYear() + 1;
  var categoryMaxAges = {'U6': 5, 'U8': 7, 'U10': 9, 'U12+': 13}; // As defined previously

  // comp_subscription_map from config.gs
  for (var categoryName in comp_subscription_map) {
    if (comp_subscription_map.hasOwnProperty(categoryName)) {
      var yearConfig = comp_subscription_map[categoryName]; // e.g., [2017, 2018] or [2014]
      var yearToTest = categoryName != 'U12+' ? yearConfig[1] : yearConfig[0];
      var maxAgeForCategory = categoryMaxAges[categoryName];

      if (maxAgeForCategory !== undefined) {
        var ageInSimulatedFutureYear = simulatedFutureYear - yearToTest;
        if (ageInSimulatedFutureYear > maxAgeForCategory) {
          allTestsPassed = false;
          Logger.log("--------------------------------------------------");
          Logger.log("FAIL: testCompSubscriptionConfiguration - Outdated category '" + categoryName + "'");
          Logger.log("  Configured Year(s): " + JSON.stringify(yearConfig) + ", Year Tested: " + yearToTest);
          Logger.log("  Simulated Future Year: " + simulatedFutureYear + ", Calculated Age: " + ageInSimulatedFutureYear);
          Logger.log("  Exceeds max age of " + maxAgeForCategory + " for this category.");
        }
      } else {
        Logger.log("WARNING: testCompSubscriptionConfiguration - Unknown category '" + categoryName + "'. Max age not defined. Skipping.");
      }
    }
  }
  Logger.log("Finished testCompSubscriptionConfiguration().");
  return allTestsPassed;
}

function testSkipassConfiguration() {
  var allTestsPassed = true;
  var simulatedFutureYear = new Date().getFullYear() + 1;

  var categoryLogic = {
    'Adulte':   { type: 'latestBirthYearFromMinAge', minCatAge: 17, configIndex: 0 },
    'Étudiant': { type: 'latestBirthYear', minCatAge: 17, configIndex: 1 },
    'Junior':   { type: 'latestBirthYear', minCatAge: 10, configIndex: 1 },
    'Enfant':   { type: 'latestBirthYear', minCatAge: 6,  configIndex: 1 },
    'Bambin':   { type: 'earliestBirthYearForMaxAge', maxCatAge: 5, configIndex: 0 },
    'Senior':   { type: 'ageRange' },
    'Vermeil':  { type: 'ageRange' }
  };

  // skipass_configuration_map from config.gs
  for (var categoryName in skipass_configuration_map) {
    if (skipass_configuration_map.hasOwnProperty(categoryName)) {
      var configValue = skipass_configuration_map[categoryName];
      var logic = categoryLogic[categoryName];

      if (logic) {
        if (logic.type === 'ageRange') {
          Logger.log("INFO: testSkipassConfiguration - Category '" + categoryName + "' is age-range defined. Config values [" + configValue.join(", ") + "] don't become outdated annually. Skipping direct year validation.");
          continue;
        }

        var expectedBirthYear;
        var actualBirthYear;

        if (logic.type === 'latestBirthYearFromMinAge' || logic.type === 'latestBirthYear') {
          expectedBirthYear = simulatedFutureYear - logic.minCatAge;
          actualBirthYear = configValue[logic.configIndex];
        } else if (logic.type === 'earliestBirthYearForMaxAge') {
          expectedBirthYear = simulatedFutureYear - logic.maxCatAge;
          actualBirthYear = configValue[logic.configIndex];
        } else { // Should not happen with current logic types
            Logger.log("WARNING: testSkipassConfiguration - Unknown logic type for category '" + categoryName + "'. Skipping.");
            continue;
        }


        if (actualBirthYear !== expectedBirthYear) {
          allTestsPassed = false;
          Logger.log("--------------------------------------------------");
          Logger.log("FAIL: testSkipassConfiguration - Outdated '" + categoryName + "' configuration.");
          Logger.log("  Type: " + logic.type + ", Config Index: " + logic.configIndex);
          Logger.log("  Configured Birth Year: " + actualBirthYear);
          Logger.log("  Simulated Future Year: " + simulatedFutureYear);
          if (logic.minCatAge) Logger.log("  Assumed Min Category Age: " + logic.minCatAge);
          if (logic.maxCatAge) Logger.log("  Assumed Max Category Age: " + logic.maxCatAge);
          Logger.log("  Expected Birth Year for " + simulatedFutureYear + ": " + expectedBirthYear);
        }
      } else {
        Logger.log("WARNING: testSkipassConfiguration - Unknown category '" + categoryName + "' in skipass_configuration_map. Test logic not defined. Skipping.");
      }
    }
  }
  Logger.log("Finished testSkipassConfiguration().");
  return allTestsPassed;
}

function testCreateSkipassMap_UsesDynamicRanges() {
  var failures = 0;

  var mockSheet = {
    getRange: function(row, col) {
      if (!this.getRangeCalls) {
        this.getRangeCalls = [];
      }
      this.getRangeCalls.push({ row: row, col: col });
      return { /* mock range object, currently no methods needed */ };
    },
    getRangeCalls: []
  };

  // This map helps convert the function call name (from createSkipassMap's source) to the base name used in skipass_configuration_map
  var getterToBaseNameMap = {
    'getSkiPassSenior': 'Senior',
    'getSkiPassSuperSenior': 'Vermeil',
    'getSkiPassAdult': 'Adulte',
    'getSkiPassStudent': 'Étudiant',
    'getSkiPassJunior': 'Junior',
    'getSkiPassKid': 'Enfant',
    'getSkiPassToddler': 'Bambin'
  };

  // The order of entries as they appear in createSkipassMap's to_return object.
  // Each entry here corresponds to one call to sheet.getRange().
  var expectedSkipassOrder = [
    { localizedKey: 'Collet Senior', getterName: 'getSkiPassSenior' },
    { localizedKey: 'Collet Vermeil', getterName: 'getSkiPassSuperSenior' },
    { localizedKey: 'Collet Adulte', getterName: 'getSkiPassAdult' },
    { localizedKey: 'Collet Étudiant', getterName: 'getSkiPassStudent' },
    { localizedKey: 'Collet Junior', getterName: 'getSkiPassJunior' },
    { localizedKey: 'Collet Enfant', getterName: 'getSkiPassKid' },
    { localizedKey: 'Collet Bambin', getterName: 'getSkiPassToddler' },
    { localizedKey: '3 Domaines Senior', getterName: 'getSkiPassSenior' },
    { localizedKey: '3 Domaines Vermeil', getterName: 'getSkiPassSuperSenior' },
    { localizedKey: '3 Domaines Adulte', getterName: 'getSkiPassAdult' },
    { localizedKey: '3 Domaines Étudiant', getterName: 'getSkiPassStudent' },
    { localizedKey: '3 Domaines Junior', getterName: 'getSkiPassJunior' },
    { localizedKey: '3 Domaines Enfant', getterName: 'getSkiPassKid' },
    { localizedKey: '3 Domaines Bambin', getterName: 'getSkiPassToddler' }
  ];

  // Call the function under test
  // Ensure skipass_configuration_map is available in the global scope for createSkipassMap
  var skiPassMapResult = createSkipassMap(mockSheet);

  // Check total number of getRange calls
  if (mockSheet.getRangeCalls.length !== expectedSkipassOrder.length) {
    failures++;
    Logger.log("--------------------------------------------------");
    Logger.log("FAIL: testCreateSkipassMap_UsesDynamicRanges - Incorrect number of getRange calls.");
    Logger.log("  Expected: " + expectedSkipassOrder.length);
    Logger.log("  Got: " + mockSheet.getRangeCalls.length);
  }

  // Check each call for correct row and column, assuming order is preserved
  for (var i = 0; i < expectedSkipassOrder.length; i++) {
    var expectedEntry = expectedSkipassOrder[i];
    var baseName = getterToBaseNameMap[expectedEntry.getterName];

    if (!baseName) {
      failures++;
      Logger.log("--------------------------------------------------");
      Logger.log("FAIL: testCreateSkipassMap_UsesDynamicRanges - Unknown getterName '" + expectedEntry.getterName + "' in test's expectedSkipassOrder for " + expectedEntry.localizedKey);
      continue;
    }

    if (!skipass_configuration_map || !skipass_configuration_map[baseName] || skipass_configuration_map[baseName].length < 4) {
      failures++;
      Logger.log("--------------------------------------------------");
      Logger.log("FAIL: testCreateSkipassMap_UsesDynamicRanges - skipass_configuration_map missing or incomplete for baseName '" + baseName + "' (derived from " + expectedEntry.localizedKey + ")");
      Logger.log("  skipass_configuration_map['" + baseName + "']: " + JSON.stringify(skipass_configuration_map ? skipass_configuration_map[baseName] : "undefined"));
      continue;
    }

    var expectedRow = skipass_configuration_map[baseName][2] + (i >= 7 ? skipass_configuration_map_3d_row_offset : 0);
    var expectedCol = skipass_configuration_map[baseName][3];

    var actualCall = mockSheet.getRangeCalls[i];

    if (!actualCall) {
      // This case should ideally be caught by the length check upfront,
      // but as a safeguard if the loop runs longer than actual calls.
      if (mockSheet.getRangeCalls.length >= expectedSkipassOrder.length) {
         // Only log if we didn't already log a length mismatch for this specific index issue
        failures++;
        Logger.log("--------------------------------------------------");
        Logger.log("FAIL: testCreateSkipassMap_UsesDynamicRanges - Missing getRange call for expected entry: " + expectedEntry.localizedKey + " at index " + i);
      }
      continue;
    }

    if (actualCall.row !== expectedRow || actualCall.col !== expectedCol) {
      failures++;
      Logger.log("--------------------------------------------------");
      Logger.log("FAIL: testCreateSkipassMap_UsesDynamicRanges - Incorrect getRange parameters for: " + expectedEntry.localizedKey + " (Base: " + baseName + ")");
      Logger.log("  Expected Row: " + expectedRow + ", Expected Col: " + expectedCol);
      Logger.log("  Actual Row: " + actualCall.row + ", Actual Col: " + actualCall.col);
      Logger.log("  (from skipass_configuration_map['" + baseName + "']: [" + skipass_configuration_map[baseName].join(", ") + "])");
    }
  }

  // Also verify that the created map has the correct keys, as a sanity check
  var numKeysInResult = Object.keys(skiPassMapResult).length;
  if (numKeysInResult !== expectedSkipassOrder.length) {
    failures++;
    Logger.log("--------------------------------------------------");
    Logger.log("FAIL: testCreateSkipassMap_UsesDynamicRanges - Incorrect number of entries in the returned skiPassMap.");
    Logger.log("  Expected: " + expectedSkipassOrder.length);
    Logger.log("  Got: " + numKeysInResult);
  }
  Logger.log("Finished testCreateSkipassMap_UsesDynamicRanges().");
  return failures === 0;
}

function testCategoryOrders() {
  function logError(cat1, cat2) {
    Logger.log("FAILURE: testCategoryOrders(" + cat1 + ", " + cat2 +")")
    return false
  }
  var cat1 = 'U6'
  var cat2 = 'U6'
  if (categoriesAscendingOrder(cat1, cat2) != 0) {
    return logError(cat1, cat2)
  }
  cat1 = 'U10'; cat2 = 'U8'
  if (categoriesAscendingOrder(cat1, cat2) != 1) {
    return logError(cat1, cat2)
  }
  cat1 = 'U6'; cat2 = 'U8'
  if (categoriesAscendingOrder(cat1, cat2) != -1) {
    return logError(cat1, cat2)
  }  
  cat1 = 'U8'; cat2 = 'U6'
  if (categoriesAscendingOrder(cat1, cat2) != 1) {
    return logError(cat1, cat2)
  }  
  cat1 = 'U12+'; cat2 = 'U10'
  if (categoriesAscendingOrder(cat1, cat2) != 1) {
    return logError(cat1, cat2)
  }
  cat1 = 'U8'; cat2 = 'U10'
  if (categoriesAscendingOrder(cat1, cat2) != -1) {
    return logError(cat1, cat2)
  }
  cat1 = 'U10'; cat2 = 'U12+'
  if (categoriesAscendingOrder(cat1, cat2) != -1) {
    return logError(cat1, cat2)
  }
  cat1 = 'U8'; cat2 = 'U12+'
  if (categoriesAscendingOrder(cat1, cat2) != -1) {
    return logError(cat1, cat2)
  }
  Logger.log("Finished testCategoryOrders().");
  return true
}

function testCreateCompSubscriptionMap() {
  var map = createCompSubscriptionMap(SpreadsheetApp.getActiveSheet())
  for (const key of ['1U6', '1U8', '1U10', '1U12+',
                     '2U6', '2U8', '2U10', '2U12+',
                     '3U6', '3U8', '3U10', '3U12+',
                     '4U6', '4U8', '4U10', '4U12+']) {
    if (!map.hasOwnProperty(key)) {
      Logger.log('FAILURE: Property not found: ' + key)
      return false
    }
  }
  Logger.log("Finished testCreateCompSubscriptionMap().");
  return true
}

function testCreateNonCompSubscriptionMap() {
  var map = createNonCompSubscriptionMap(SpreadsheetApp.getActiveSheet())
  for (const key of ['Rider', '1er enfant',
                     '2ème enfant', '3ème enfant',
                     '4ème enfant']) {
    if (!map.hasOwnProperty(key)) {
      Logger.log('FAILURE: Property not found: ' + key)
      return false
    }
  }
  Logger.log("Finished testCreateNonCompSubscriptionMap().");
  return true
}

function areArraysEqual(arr1, arr2) {
  if (arr1.length !== arr2.length) { return false }
  for (let i = 0; i < arr1.length; i++) {
    if (arr1[i] !== arr2[i]) {
      return false;
    } 
  }
  return true
}

function testIsLevel() {
  var existing_levels = {
    //                  NotAdjusted | NotDefined | Defined | Comp |  NotComp | Rider | RecreationalNonRider | LicenseOnly
    "":                [false,        true,        false,    false,  false,    false,  false,                 false],
    "Licence seule":   [false,        false,       true,     false,  false,    false,  false,                 true ],
    "Non déterminé":   [false,        false,       true,     false,  true,     false,  true,                  false],
    "Compétiteur":     [false,        false,       true,     true,   false,    false,  false,                 false],
    "Débutant/Ourson": [false,        false,       true,     false,  true,     false,  true,                  false],
    "Flocon":          [false,        false,       true,     false,  true,     false,  true,                  false],
    "Étoile 1":        [false,        false,       true,     false,  true,     false,  true,                  false],
    "Étoile 2":        [false,        false,       true,     false,  true,     false,  true,                  false],
    "Étoile 3":        [false,        false,       true,     false,  true,     false,  true,                  false],
    "Bronze":          [false,        false,       true,     false,  true,     false,  true,                  false],
    "Argent":          [false,        false,       true,     false,  true,     false,  true,                  false],
    "Or":              [false,        false,       true,     false,  true,     false,  true,                  false],
    "Ski/Fun":         [false,        false,       true,     false,  true,     false,  true,                  false],
    "Rider":           [false,        false,       true,     false,  true,     true,   false,                 false],
    "Snow Découverte": [false,        false,       true,     false,  true,     false,  true,                  false],
    "Snow 1":          [false,        false,       true,     false,  true,     false,  true,                  false],
    "Snow 2":          [false,        false,       true,     false,  true,     false,  true,                  false],
    "Snow 3":          [false,        false,       true,     false,  true,     false,  true,                  false],
    "Snow Expert":     [false,        false,       true,     false,  true,     false,  true,                  false],
  }

  for (const [key, values] of Object.entries(existing_levels)) {
      var results = [isLevelNotAdjusted(key),
                     isLevelNotDefined(key),
                     isLevelDefined(key),
                     isLevelComp(key),
                     isLevelNotComp(key),
                     isLevelRider(key),
                     isLevelRecreationalNonRider(key),
                     isLevelLicenseOnly(key)]
      // Logger.log('INFO: [' + key + ']: ' + arrayToRawString(results))
      if (! areArraysEqual(values, results)) {
        Logger.log('FAILURE: for existing_levels[' + key + ']. Got:' + arrayToRawString(results) + ', expected: ' + arrayToRawString(values))
        return false
      }
  }
  var existing_levels_not_set = {}
  for (const [key, value_] of Object.entries(existing_levels)) {
    if (key == '') {
      continue
    }
    entry = "⚠️ " + key
    existing_levels_not_set[entry] = true
  }
  for (const [key, value] of Object.entries(existing_levels_not_set)) {
    var result = isLevelNotAdjusted(key)
    if (result != value) {
      Logger.log('FAILURE: for existing_levels_not_set[' + key + ']. Got:' + result + ', expected: ' + value)
      return false      
    }
  }
  Logger.log("Finished testIsLevel().");
  return true
}

function testIsLicense() {
  var existing_licenses = {
    //                          Defined | NoLicense | NotDefined | NonCompJunior | NonCompAdult | NonCompFamily | NonComp | Comp | CompAdult | CompJunior | Exec
    'Z@%!':                    [false,    false,      true,        false,          false,         false,          false,    false, false,      false,       false],
    'Aucune':                  [false,    true,       true,        false,          false,         false,          false,    false, false,      false,       false],
    'CN Jeune (Loisir)':       [true,     false,      false,       true,           false,         false,          true,     false, false,      false,       false],
    'CN Adulte (Loisir)':      [true,     false,      false,       false,          true,          false,          true,     false, false,      false,       false],
    'CN Famille (Loisir)':     [true,     false,      false,       false,          false,         true,           true,     false, false,      false,       false],
    'CN Dirigeant':            [true,     false,      false,       false,          false,         false,          false,    false, false,      false,       true ],
    'CN Jeune (Compétition)':  [true,     false,      false,       false,          false,         false,          false,    true,  false,      true,        false],
    'CN Adulte (Compétition)': [true,     false,      false,       false,          false,         false,          false,    true,  true,       false,       false],
  }
  for (const [key, values] of Object.entries(existing_licenses)) {
    var results = [isLicenseDefined(key),
                   isLicenseNoLicense(key),
                   isLicenseNotDefined(key),
                   isLicenseNonCompJunior(key),
                   isLicenseNonCompAdult(key),
                   isLicenseNonCompFamily(key),
                   isLicenseNonComp(key),
                   isLicenseComp(key),
                   isLicenseCompAdult(key),
                   isLicenseCompJunior(key),
                   isLicenseExec(key)]
    // Logger.log('INFO: [' + key + ']: ' + arrayToRawString(results))
    if (! areArraysEqual(values, results)) {
      Logger.log('FAILURE: for existing_licences[' + key + ']. Got:' + arrayToRawString(results) + ', expected: ' + arrayToRawString(values))
      return false
    }
  }
  Logger.log("Finished testIsLicense().");
  return true
}

function testPhoneNumberValidation() {
  const good_phone_number = "01 01 01 01 01"
  if (!validatePhoneNumber(good_phone_number) == true) {
    Logger.log('FAILURE: failling good phone number validation: ' + good_phone_number)
    return false
  }
  const bad_phone_number = [
   "06 87 52 23",
   "06 87 52 23 73 ",
   "06 87 52 23 733",
   "A6 87 52 23 73",
  ]
  for (const phone of bad_phone_number) {
    if (validatePhoneNumber(phone)) {
      Logger.log('FAILURE: failling bad phone number validation: ' + phone)
      return false
    }
  }
  Logger.log("Finished testPhoneNumberValidation().");
  return true  
}

const originalDateNow = Date
function startMockDate(date) {
  date_object = new Date(date)
  Date = function() {
    return date_object
  }
}
function stopMockDate() {
  Date = originalDateNow
}

function testAgeVerification() {
  var test_values_1 = {
    //         DoB                    | To Test Against       | Method                               | Expected results
    'Test 1a': [new Date("12/31/2006"), new Date("12/31/2006"), ageVerificationBornBeforeDateIncluded, true           ],
    'Test 1b': [new Date("01/01/2006"), new Date("12/31/2006"), ageVerificationBornBeforeDateIncluded, true           ],
    'Test 1c': [new Date("01/01/2007"), new Date("12/31/2006"), ageVerificationBornBeforeDateIncluded, false          ],

    'Test 1d': [new Date("01/01/2020"), new Date("01/01/2020"), ageVerificationBornAfterDateIncluded,  true           ],
    'Test 1e': [new Date("01/01/2021"), new Date("01/01/2020"), ageVerificationBornAfterDateIncluded,  true           ],
    'Test 1f': [new Date("12/31/2019"), new Date("01/01/2020"), ageVerificationBornAfterDateIncluded,  false          ],

    'Test 1g': [new Date("12/31/2015"), "2016",                 ageVerificationBornBeforeYearIncluded, true           ],
    'Test 1h': [new Date("12/31/2016"), "2016",                 ageVerificationBornBeforeYearIncluded, true           ],
    'Test 1i': [new Date("1/1/2017"),   "2016",                 ageVerificationBornBeforeYearIncluded, false          ],
  }
  for (const [key, values] of Object.entries(test_values_1)) {
    var dob = values[0]
    var date = values[1]
    var method = values[2]
    var result = values[3]
    if (method(dob, date) != result) {
      Logger.log('FAILURE: ' + key + ": " + dob + " and " + date)
      return false
    }
  }

  var test_values_2 = {
    //         DoB                    | Date 1                | Date 2                | method                                 | Expected results
    'Test 2a': [new Date("1/1/2005"),   new Date("1/1/2005"),   new Date("12/31/2005"), ageVerificationBornBetweenDatesIncluded, true            ],
    'Test 2b': [new Date("12/31/2005"), new Date("1/1/2005"),   new Date("12/31/2005"), ageVerificationBornBetweenDatesIncluded, true            ],
    'Test 2c': [new Date("1/1/2006"),   new Date("1/1/2005"),   new Date("12/31/2005"), ageVerificationBornBetweenDatesIncluded, false           ],
    'Test 2d': [new Date("12/31/2004"), new Date("1/1/2005"),   new Date("12/31/2005"), ageVerificationBornBetweenDatesIncluded, false           ],
    //         Dob                    | Year 1                | Year 2                | Method                                 | Expected results
    'Test 2e': [new Date("1/1/2005"),   2005,                   2006,                   ageVerificationBornBetweenYearsIncluded, true            ],
    'Test 2f': [new Date("12/31/2005"), 2005,                   2006,                   ageVerificationBornBetweenYearsIncluded, true            ],
    'Test 2g': [new Date("12/31/2004"), 2005,                   2006,                   ageVerificationBornBetweenYearsIncluded, false           ],
    'Test 2h': [new Date("1/1/2007"),   2005,                   2006,                   ageVerificationBornBetweenYearsIncluded, false           ],
  }
  for (const [key, values] of Object.entries(test_values_2)) {
    var dob = values[0]
    var date1 = values[1]
    var date2 = values[2]
    var method = values[3]
    var result = values[4]
    if (method(dob, date1, date2) != result) {
      Logger.log('FAILURE: ' + key + ": " + dob + " and " + date1 + ", " + date2)
      return false
    }
  }

  var test_values_3 = {
    //         Dob                   | Age
    'Test 3a': [new Date("10/8/2011"), 14 ], // 14 today
    'Test 3b': [new Date("10/7/2011"), 14 ], // 14 yesterday
    'Test 3c': [new Date("09/8/2011"), 14 ], // 14 for a month
    'Test 3d': [new Date("11/8/2011"), 13 ], // 14 in one month
    'Test 3e': [new Date("10/9/2011"), 13 ]  // 14 in one day
  }

  startMockDate("10/8/2025")
  for (const [key, values] of Object.entries(test_values_3)) {
    var dob = values[0]
    var age = values[1]
    var res = ageFromDoB(dob)
    if (res != age) {
      Logger.log('FAILURE: ' + key + ": " + dob + " and " + age + ", res:" + res)
      stopMockDate()
      return false
    }
  }
  stopMockDate()

  var test_values_4 = {
    //         DoB                   | Age | method                           | expected              
    'Test 4a': [new Date("10/8/2011"), 14,   ageVerificationStrictlyOldOrOlder, true    ],
    'Test 4b': [new Date("10/9/2011"), 13,   ageVerificationStrictlyOldOrOlder, true    ],
    'Test 4c': [new Date("10/9/2011"), 14,   ageVerificationStrictlyOldOrOlder, false   ],
    'Test 4d': [new Date("10/9/2011"), 14,   ageVerificationStrictlyYounger,    true    ],
    'Test 4e': [new Date("10/8/2011"), 13,   ageVerificationStrictlyYounger,    false   ],
    'Test 4f': [new Date("10/9/2011"), 14,   ageVerificationStrictlyYounger,    true    ]

  }
  startMockDate("10/8/2025")
  for (const [key, values] of Object.entries(test_values_4)) {
    var dob = values[0]
    var age = values[1]
    var method = values[2]
    var expected = values[3]
    var res = method(dob, age)
    if (res != expected) {
      Logger.log('FAILURE: ' + key + ": " + dob + " and " + age + " res=" + res)
      stopMockDate()
      return false
    }
  }
  stopMockDate()

  var test_values_5 = {
    'Test 5a': [new Date("10/8/2025"), true ],
    'Test 5b': [new Date("10/7/2025"), true ],
    'Test 5c': [new Date("10/9/2025"), false]

  }
  startMockDate("10/8/2043")
  for (const [key, values] of Object.entries(test_values_5)) {
    var dob = values[0]
    var expected = values[1]
    var res = isAdult(dob)
    if (res != expected) {
      Logger.log('FAILURE: ' + key + ": " + dob + " expected=" + expected + " res=" + res)
      stopMockDate()
      return false      
    }
  }
  stopMockDate()

  var test_values_6 = {
    //         dob                   | expected
    'Test 6a': [new Date("10/8/2025"), Number(2025)    ],
    'Test 6b': ["1/1/1901",            Number(1901)    ],
    // NaN is the only value that isn't equal to itself.
    // 'Test 6c': ["",                    NaN             ]
  }
  for (const [key, values] of Object.entries(test_values_6)) {
    var dob = values[0]
    var expected = values[1]
    var res = getDoBYear(dob)
    if (res != expected) {
      Logger.log('FAILURE: ' + key + ": " + dob + " expected=" + expected + " res=" + res)
      return false
    }
  }

  var test_values_7 = {
    //          DoB                   | Age1 | Age2 | method                      | expected
    'Test 7a': [new Date("10/9/2015"),  10,    11,    ageVerificationRangeIncluded, true    ],
    'Test 7b': [new Date("10/9/2015"),  09,    10,    ageVerificationRangeIncluded, true    ],
    'Test 7c': [new Date("10/9/2015"),  10,    10,    ageVerificationRangeIncluded, true    ],
    'Test 7d': [new Date("10/9/2015"),  11,    10,    ageVerificationRangeIncluded, false   ],
    'Test 7e': [new Date("10/9/2015"),  11,    15,    ageVerificationRangeIncluded, false   ],
    'Test 7f': [new Date("10/10/2015"), 10,    11,    ageVerificationRangeIncluded, false   ],
    'Test 7g': [new Date("10/9/2014"),  11,    11,    ageVerificationRangeIncluded, true    ],
    'Test 7h': [new Date("10/9/2014"),  10,    12,    ageVerificationRangeIncluded, true    ],
    'Test 7i': [new Date("10/9/2015"),  11,    12,    ageVerificationRangeIncluded, false   ],
  }
  startMockDate("10/9/2025")
  for (const [key, values] of Object.entries(test_values_7)) {
    var dob = values[0]
    var age1 = values[1]
    var age2 = values[2]
    var method = values[3]
    var expected = values[4]
    var res = method(dob, age1, age2)
    if (res != expected) {
      Logger.log('FAILURE: ' + key + ": " + dob + " age1=" + age1 + " age2=" + age2 + " expected=" + expected)
      stopMockDate()
      return false      
    }
  }
  stopMockDate()

  Logger.log("Finished testAgeVerification")
  return true
}

function testSkipassProperties() {
  var test_values = {
    //         Input                | Method under test | Expected result
    'Test a': ["Collet Senior",       isSkipassStudent,   false          ],
    'Test b': ["Collet Étudiant",     isSkipassStudent,   true           ],
    'Test c': ["3 Domaines Étudiant", isSkipassStudent,   true           ],
    'Test d': ["2 Alpes Étudiant",    isSkipassStudent,   false          ],

    'Test e': ["Collet Senior",       isSkipassAdult,     false          ],
    'Test f': ["Collet Adulte",       isSkipassAdult,     true           ],
    'Test g': ["3 Domaines Adulte",   isSkipassAdult,     true           ],
    'Test h': ["2 Alpes Adulte",      isSkipassAdult,     false          ], 

    'Test i': ["Collet Senior",       isSkiPassToddler,   false          ],
    'Test j': ["Collet Bambin",       isSkiPassToddler,   true           ],
    'Test k': ["3 Domaines Bambin",   isSkiPassToddler,   true           ],
    'Test l': ["2 Alpes Bambin",      isSkiPassToddler,   false          ]
  }
  for (const [key, values] of Object.entries(test_values)) {
    var input = values[0]
    var method = values[1]
    var result = values[2]
    if (method(input) != result) {
      Logger.log('FAILURE: ' + key + ": " + input + ": " + result)
      return false
    }
  }

  var pairs = getSkiPassesPairs()
  var expected = [["Collet Senior","3 Domaines Senior"],
                  ["Collet Vermeil","3 Domaines Vermeil"],
                  ["Collet Adulte","3 Domaines Adulte"],
                  ["Collet Étudiant","3 Domaines Étudiant"],
                  ["Collet Junior","3 Domaines Junior"],
                  ["Collet Enfant","3 Domaines Enfant"],
                  ["Collet Bambin","3 Domaines Bambin"]]
  if (pairs.length != expected.length) {
      Logger.log('FAILURE: ' + pairs.length + ": " + expected.length)
      return false
  }
  for (let i = 0; i < pairs.length; i++) {
    if (pairs[i][0] != expected[i][0] || pairs[i][1] != expected[i][1]) {
      Logger.log('FAILURE: ' + pairs[i] + ": " + expected[i])
      return false
    }
  }
  Logger.log("Finished testSkiPassProperties")
  return true
}

function testMiscUtility() {
  var test_values_nth_string = {
    'Test 1': [-1,  "??ème"],
    'Test 2': [0,   "??ème"],
    'Test 3': [1,   "1er"  ],
    'Test 4': [2,   "2ème" ],
    'Test 5': [3,   "3ème" ],
    'Test 6': [4,   "4ème" ],
    'Test 7': [5,   "5ème" ],
    'Test 8': [6,   "??ème"],
    'Test 9': ["a", "??ème"]
  }
  for (const [key, values] of Object.entries(test_values_nth_string)) {
    if (nthString(values[0]) != values[1]) {
      Logger.log('FAILURE: ' + key + ": " + nthString(values[0]) + ", " + values[1])
      return false
    }
  }

  if (comp_subscription_categories[0] != 'U6' ||
      comp_subscription_categories[1] != 'U8' ||
      comp_subscription_categories[2] != 'U10' ||
      comp_subscription_categories[3] != 'U12+') {
      Logger.log('FAILURE: ' + comp_subscription_categories)
      return false
  }

  Logger.log("Finished testMiscUtility")
  return true
}

function RUN_ALL_TESTS() {
  Logger.log("Starting all invoice tests...");
  var failedSuites = [];

  if(!testMiscUtility()) {
    failedSuites.push('testMiscUtility')
  }
  if (!testSkipassProperties()) {
    failedSuites.push('testSkipassProperties')
  }
  if (!testAgeVerification()) {
    failedSuites.push('testAgeVerification')
  }
  if (!testPhoneNumberValidation()) {
    failedSuites.push('testPhoneNumberValidation')
  }
  if (!testIsLicense()) {
    failedSuites.push('testIsLicense')
  }
  if (!testIsLevel()) {
    failedSuites.push("testIsLevel");
  }
  if (!testCategoryOrders()) {
    failedSuites.push("testCategoryOrders");
  }
  if (!testCreateCompSubscriptionMap()) {
    failedSuites.push("testCreateCompSubscriptionMap");
  }
  if (!testCreateNonCompSubscriptionMap()) {
    failedSuites.push("testCreateNonCompSubscriptionMap");
  }
  if (!testNormalizeName()) {
    failedSuites.push("testNormalizeName");
  }
  if (!testPlural()) {
    failedSuites.push("testPlural");
  }
  if (!testGetDoBYear()) {
    failedSuites.push("testGetDoBYear");
  }
  if (!testIsLicenseDefined()) {
    failedSuites.push("testIsLicenseDefined");
  }
  if (!testAgeFromDoB()) {
    failedSuites.push("testAgeFromDoB");
  }
  if (!testFormatPhoneNumberString()) {
    failedSuites.push("testFormatPhoneNumberString");
  }
  if (!testSeasonConfiguration()) {
    failedSuites.push("testSeasonConfiguration");
  }
  if (!testLicensesConfiguration()) {
    failedSuites.push("testLicensesConfiguration");
  }
  if (!testCompSubscriptionConfiguration()) {
    failedSuites.push("testCompSubscriptionConfiguration");
  }
  // if (!testSkipassConfiguration()) {
  //   failedSuites.push("testSkipassConfiguration");
  // }
  if (!testCreateSkipassMap_UsesDynamicRanges()) {
    failedSuites.push("testCreateSkipassMap_UsesDynamicRanges");
  }
  if (failedSuites.length === 0) {
    Logger.log("Summary: ✅✅✅ All test suites passed successfully!");
  } else {
    Logger.log("Summary: ⚠️⚠️⚠️ Failures detected in the following test suite(s): " + failedSuites.join(", ") + ".");
  }
}
