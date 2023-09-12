function testRemoveDuplicates() {
  /*
  // 1. create an array with duplicates
  // 2. call the function
  // 3. check if the result is the expected result
  */
  const inputArray = [
    [1, 2],
    [2, 1],
    [3, 1],
    [1, 2],
  ];
  const outputArray = [
    [1, 2],
    [2, 1],
    [3, 1],
  ];
  const result = removeDuplicates(inputArray);

  if (arraysAreEqual(result, outputArray)) {
    console.log("testRemoveDuplicates passed!");
  } else {
    console.error("testRemoveDuplicates failed!");
  }
}
function removeDuplicates(arr) {
  /*
  // 1. create an object to store unique values
  // 2. loop through the array
  // 3. check if the element is in the object
  // 4. if yes, continue
  // 5. if no, add the element to the object

  :param arr: the array

  :return: the array without duplicates
  */
  return arr.filter((item, index) => {
    // Use the findIndex method to check for the first occurrence of the item
    const firstIndex = arr.findIndex((elem) => {
      return JSON.stringify(elem) === JSON.stringify(item);
    });
    return firstIndex === index;
  });
}

function testfindValuesNotInArray() {
  /*
  // 1. create two arrays
  // 2. call the function
  // 3. check if the result is the expected result
  */
  const arr1 = [
    [1, 2],
    [3, 4],
    [3, 2],
  ];
  const arr2 = [
    [1, 2],
    [5, 6],
    [7, 8],
    [3, 2],
  ];

  const result = findValuesNotInArray(arr1, arr2);

  outputArray = [
    [5, 6],
    [7, 8],
  ];
  if (arraysAreEqual(result, outputArray)) {
    console.log("testfindValuesNotInArray passed!");
  } else {
    console.error("testfindValuesNotInArray failed!");
  }
}
function findValuesNotInArray(arr1, arr2) {
  /*
  // 1. loop through the second array
  // 2. check if the element is in the first array
  // 3. if yes, continue
  // 4. if no, add the element to the result array

  :param arr1: the first array
  :param arr2: the second array

  :return: the elements of the second array that are not in the first array
  */
  const result = [];

  for (const item2 of arr2) {
    let found = false;

    for (const item1 of arr1) {
      if (arraysAreEqual(item1, item2)) {
        found = true;
        break;
      }
    }

    if (!found) {
      result.push(item2);
    }
  }

  return result;
}

// Helper function to check if two arrays are deeply equal
function arraysAreEqual(arr1, arr2) {
  /*
  // 1. check if the length of the arrays are equal
  // 2. check if each element of the arrays are equal
  // 3. if yes, return true
  // 4. if no, return false

  :param arr1: the first array
  :param arr2: the second array

  :return: true if the arrays are equal, false otherwise
  */
  if (arr1.length !== arr2.length) {
    return false;
  }

  for (let i = 0; i < arr1.length; i++) {
    if (JSON.stringify(arr1[i]) !== JSON.stringify(arr2[i])) {
      return false;
    }
  }

  return true;
}

testRemoveDuplicates();
testfindValuesNotInArray();
