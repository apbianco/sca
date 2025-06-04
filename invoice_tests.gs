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

function runInvoiceTests() {
  Logger.log("Starting all invoice tests...");
  var failedSuites = [];

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
  if (failedSuites.length === 0) {
    Logger.log("Summary: All test suites passed successfully!");
  } else {
    Logger.log("Summary: Failures detected in the following test suite(s): " + failedSuites.join(", ") + ".");
  }
}
