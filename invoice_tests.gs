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
      description: "Multi-word input",
      number: 3,
      message: "apple tree",
      expected: "apples trees"
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
      description: "Input with only spaces, singular",
      number: 1,
      message: " ",
      expected: " "
    },
    {
      description: "Input with only spaces, plural",
      number: 2,
      message: " ",
      expected: " "
    }
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
    expected: Math.abs(year_future - 1970)
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
  return true
}

function testIsLevel() {
  function areArraysEqual(arr1, arr2) {
    if (arr1.length !== arr2.length) { return false }
    for (let i = 0; i < arr1.length; i++) {
      if (arr1[i] !== arr2[i]) {
        return false;
      } 
    }
    return true
  }
  var existing_levels = {
    //                  NotAdjusted | NotDefined | Defined | Comp |  NotComp | Rider | RecreationalNonRider
    "":                [false,        true,        false,    false,  false,    false,  false],
    "Pas Concerné":    [false,        true,        false,    false,  false,    false,  false],
    "Non déterminé":   [false,        false,       true,     false,  true,     false,  true],
    "Compétiteur":     [false,        false,       true,     true,   false,    false,  false],
    "Débutant/Ourson": [false,        false,       true,     false,  true,     false,  true],
    "Flocon":          [false,        false,       true,     false,  true,     false,  true],
    "Étoile 1":        [false,        false,       true,     false,  true,     false,  true],
    "Étoile 2":        [false,        false,       true,     false,  true,     false,  true],
    "Étoile 3":        [false,        false,       true,     false,  true,     false,  true],
    "Bronze":          [false,        false,       true,     false,  true,     false,  true],
    "Argent":          [false,        false,       true,     false,  true,     false,  true],
    "Or":              [false,        false,       true,     false,  true,     false,  true],
    "Ski/Fun":         [false,        false,       true,     false,  true,     false,  true],
    "Rider":           [false,        false,       true,     false,  true,     true,   false],
    "Snow Découverte": [false,        false,       true,     false,  true,     false,  true],
    "Snow 1":          [false,        false,       true,     false,  true,     false,  true],
    "Snow 2":          [false,        false,       true,     false,  true,     false,  true],
    "Snow 3":          [false,        false,       true,     false,  true,     false,  true],
    "Snow Expert":     [false,        false,       true,     false,  true,     false,  true],
  }

  for (const [key, values] of Object.entries(existing_levels)) {
      var results = [isLevelNotAdjusted(key),
                     isLevelNotDefined(key),
                     isLevelDefined(key),
                     isLevelComp(key),
                     isLevelNotComp(key),
                     isLevelRider(key),
                     isLevelRecreationalNonRider(key)]
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
  return true
}

function RUN_ALL_TESTS() {
  Logger.log("Starting all invoice tests...");
  var failedSuites = [];

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
