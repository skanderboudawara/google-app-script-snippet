function testRemoveDuplicates() {
  const inputArray = [[1, 2], [2, 1], [3, 1], [1, 2]];
  const outputArray = [[1, 2], [2, 1], [3, 1]];
  const result = removeDuplicates(inputArray);

  if (arraysAreEqual(result, outputArray)) {
    console.log("testRemoveDuplicates passed!");
  } else {
    console.error("testRemoveDuplicates failed!");
  }
}
function removeDuplicates(arr) {
  // Create an object to store unique values
  return arr.filter((item, index) => {
    // Use the findIndex method to check for the first occurrence of the item
    const firstIndex = arr.findIndex((elem) => {
      return JSON.stringify(elem) === JSON.stringify(item);
    });
    return firstIndex === index;
  });
}

function testfindValuesNotInArray() {
  const arr1 = [[1, 2], [3, 4], [3, 2]];
  const arr2 = [[1, 2], [5, 6], [7, 8], [3, 2]];

  const result = findValuesNotInArray(arr1, arr2);

  outputArray = [[5, 6], [7, 8]]
  if (arraysAreEqual(result, outputArray)) {
    console.log("testfindValuesNotInArray passed!");
  } else {
    console.error("testfindValuesNotInArray failed!");
  }
}
function findValuesNotInArray(arr1, arr2) {
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
