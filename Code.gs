const SS = SpreadsheetApp.getActiveSpreadsheet();

/* =========================
   WEB APP
========================= */

function doGet() {
  return HtmlService
    .createHtmlOutputFromFile('Index')
    .setTitle('🍜 กินอะไรดี?')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* =========================
   BASIC HELPERS
========================= */

function getSheet(name) {
  const sheet = SS.getSheetByName(name);

  if (!sheet) {
    throw new Error('ไม่พบ Sheet: ' + name);
  }

  return sheet;
}

function normalizeValue(value) {
  if (value instanceof Date) {
    return Utilities.formatDate(
      value,
      Session.getScriptTimeZone(),
      'yyyy-MM-dd HH:mm:ss'
    );
  }

  return value;
}

function isTrue(value) {
  if (value === true) return true;

  const text = String(value).trim().toLowerCase();

  return (
    text === 'true' ||
    text === '1' ||
    text === 'yes' ||
    text === 'ใช่'
  );
}

function getData(name) {
  const sheet = getSheet(name);
  const values = sheet.getDataRange().getValues();

  if (values.length <= 1) {
    return [];
  }

  const headers = values[0];

  return values.slice(1).map(function(row) {
    const obj = {};

    headers.forEach(function(header, i) {
      obj[String(header).trim()] = normalizeValue(row[i]);
    });

    return obj;
  });
}

function makeId(prefix) {
  return (
    prefix +
    Utilities.getUuid()
      .replace(/-/g, '')
      .substring(0, 12)
  );
}

/* =========================
   APP DATA
========================= */

function getAppData() {
  return {
    foods: getData('Foods'),
    categories: getData('Categories'),
    ratings: getData('Ratings'),
    history: getData('History'),
    people: getData('People'),
    ingredients: getData('Ingredients')
  };
}

/* =========================
   CATEGORIES
========================= */

function getCategories() {
  const categories = getData('Categories');

  return {
    methods: categories.filter(function(item) {
      return (
        isTrue(item.Active) &&
        String(item['Category Type']).trim() === 'วิธีทาน'
      );
    }),

    foodTypes: categories.filter(function(item) {
      return (
        isTrue(item.Active) &&
        String(item['Category Type']).trim() === 'ประเภทอาหาร'
      );
    })
  };
}

/* =========================
   FOOD
========================= */

function addFood(food) {
  if (!food) {
    throw new Error('ไม่มีข้อมูลเมนู');
  }

  const sheet = getSheet('Foods');
  const id = makeId('F');

  const method = String(food.eatMethod || '').trim();
  const isHome = method === 'ทำกินเองที่บ้าน';

  sheet.appendRow([
    id,                              // A ID
    food.name || '',                // B ชื่อเมนู
    method,                          // C วิธีทาน
    food.foodType || '',             // D ประเภทอาหาร
    isHome ? '' : (food.shop || ''), // E ร้าน
    food.price || '',                // F ราคา
    food.note || '',                 // G หมายเหตุ
    true,                            // H Active
    new Date(),                      // I วันที่เพิ่ม
    Boolean(food.favorite),          // J Favorite
    isHome ? (food.servings || '') : '', // K จำนวนที่ทำได้
    isHome ? (food.recipe || '') : '',   // L วิธีทำ
    isHome ? (food.recipeUrl || '') : '', // M Recipe URL
    isHome ? '' : (food.mapsUrl || '')     // N Google Maps URL
  ]);

  if (isHome) {
    saveIngredients(id, food.ingredients || []);
  }

  return {
    success: true,
    id: id
  };
}

function updateFood(food) {
  if (!food || !food.id) {
    return {
      success: false,
      message: 'ไม่พบ ID ของเมนู'
    };
  }

  const sheet = getSheet('Foods');
  const values = sheet.getDataRange().getValues();

  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(food.id)) {

      const method = String(food.eatMethod || '').trim();
      const isHome = method === 'ทำกินเองที่บ้าน';

      sheet.getRange(i + 1, 2, 1, 13).setValues([[
        food.name || '',                         // B
        method,                                  // C
        food.foodType || '',                     // D
        isHome ? '' : (food.shop || ''),         // E
        food.price || '',                        // F
        food.note || '',                         // G
        values[i][7],                            // H Active
        values[i][8],                            // I วันที่เพิ่ม
        Boolean(food.favorite),                  // J
        isHome ? (food.servings || '') : '',     // K
        isHome ? (food.recipe || '') : '',       // L
        isHome ? (food.recipeUrl || '') : '',    // M
        isHome ? '' : (food.mapsUrl || '')       // N
      ]]);

      if (isHome) {
        saveIngredients(food.id, food.ingredients || []);
      } else {
        deleteIngredients(food.id);
      }

      return {
        success: true,
        message: 'บันทึกการแก้ไขแล้ว'
      };
    }
  }

  return {
    success: false,
    message: 'ไม่พบเมนู'
  };
}

function disableFood(foodId) {
  return setFoodActive(foodId, false);
}

function enableFood(foodId) {
  return setFoodActive(foodId, true);
}

function setFoodActive(foodId, active) {
  const sheet = getSheet('Foods');
  const values = sheet.getDataRange().getValues();

  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(foodId)) {

      sheet.getRange(i + 1, 8).setValue(active);

      return {
        success: true,
        active: active
      };
    }
  }

  return {
    success: false,
    message: 'ไม่พบเมนู'
  };
}

function toggleFavorite(foodId) {
  const sheet = getSheet('Foods');
  const values = sheet.getDataRange().getValues();

  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(foodId)) {

      const current = isTrue(values[i][9]);
      const next = !current;

      sheet.getRange(i + 1, 10).setValue(next);

      return {
        success: true,
        favorite: next
      };
    }
  }

  return {
    success: false,
    message: 'ไม่พบเมนู'
  };
}

/* =========================
   INGREDIENTS
========================= */

function saveIngredients(foodId, ingredients) {
  deleteIngredients(foodId);

  if (!ingredients || !ingredients.length) {
    return;
  }

  const sheet = getSheet('Ingredients');

  const rows = [];

  ingredients.forEach(function(item) {

    if (!item || !String(item.name || '').trim()) {
      return;
    }

    rows.push([
      makeId('I'),
      foodId,
      item.name || '',
      item.amount || '',
      item.unit || '',
      item.note || ''
    ]);
  });

  if (rows.length) {
    sheet
      .getRange(
        sheet.getLastRow() + 1,
        1,
        rows.length,
        6
      )
      .setValues(rows);
  }
}

function deleteIngredients(foodId) {
  const sheet = getSheet('Ingredients');
  const values = sheet.getDataRange().getValues();

  for (let i = values.length - 1; i >= 1; i--) {

    if (String(values[i][1]) === String(foodId)) {
      sheet.deleteRow(i + 1);
    }
  }
}

function getIngredients(foodId) {
  return getData('Ingredients').filter(function(item) {
    return String(item.FoodID) === String(foodId);
  });
}

/* =========================
   RATINGS
========================= */

function addRating(rating) {
  if (!rating) {
    throw new Error('ไม่มีข้อมูลคะแนน');
  }

  return updateRating(
    rating.foodId,
    rating.person,
    rating.rating,
    rating.comment
  );
}

function updateRating(foodId, person, rating, comment) {

  const lock = LockService.getScriptLock();

  lock.waitLock(10000);

  try {

    const sheet = getSheet('Ratings');
    const values = sheet.getDataRange().getValues();

    const score = Number(rating);

    if (!foodId) {
      throw new Error('ไม่พบเมนู');
    }

    if (!person) {
      throw new Error('กรุณาเลือกผู้ให้คะแนน');
    }

    if (score < 1 || score > 5) {
      throw new Error('คะแนนต้องอยู่ระหว่าง 1-5');
    }

    for (let i = 1; i < values.length; i++) {

      if (
        String(values[i][1]) === String(foodId) &&
        String(values[i][2]).trim() === String(person).trim()
      ) {

        sheet.getRange(i + 1, 4).setValue(score);
        sheet.getRange(i + 1, 5).setValue(comment || '');
        sheet.getRange(i + 1, 6).setValue(new Date());

        return {
          success: true,
          updated: true
        };
      }
    }

    sheet.appendRow([
      makeId('R'),
      foodId,
      person,
      score,
      comment || '',
      new Date()
    ]);

    return {
      success: true,
      updated: false
    };

  } finally {
    lock.releaseLock();
  }
}

/* =========================
   HISTORY
========================= */

function addHistory(foodId, method, note) {

  const sheet = getSheet('History');

  const id = makeId('H');

  sheet.appendRow([
    id,
    foodId,
    new Date(),
    method || 'สุ่ม',
    note || ''
  ]);

  return {
    success: true,
    id: id
  };
}

/* =========================
   FOOD DETAIL
========================= */

function getFoodDetail(foodId) {

  const foods = getData('Foods');
  const ratings = getData('Ratings');
  const ingredients = getData('Ingredients');

  const food = foods.find(function(item) {
    return String(item.ID) === String(foodId);
  });

  if (!food) {
    return null;
  }

  const foodRatings = ratings.filter(function(item) {
    return String(item.FoodID) === String(foodId);
  });

  const ratingValues = foodRatings
    .map(function(item) {
      return Number(item.Rating);
    })
    .filter(function(score) {
      return score >= 1 && score <= 5;
    });

  const average = ratingValues.length
    ? ratingValues.reduce(function(a, b) {
        return a + b;
      }, 0) / ratingValues.length
    : 0;

  return {
    food: food,
    ratings: foodRatings,
    ingredients: ingredients.filter(function(item) {
      return String(item.FoodID) === String(foodId);
    }),
    average: average
  };
}

/* =========================
   RANDOM / SMART RANDOM
========================= */

function randomFoods(filters) {

  filters = filters || {};

  const foods = getData('Foods')
    .filter(function(food) {
      return isTrue(food.Active);
    });

  const ratings = getData('Ratings');
  const history = getData('History');

  let result = foods.slice();

  /* วิธีทาน */
  if (filters.method) {
    result = result.filter(function(food) {
      return String(food['วิธีทาน']).trim() === String(filters.method).trim();
    });
  }

  /* ประเภทอาหาร */
  if (filters.foodType) {
    result = result.filter(function(food) {
      return String(food['ประเภทอาหาร']).trim() === String(filters.foodType).trim();
    });
  }

  /* Favorite */
  if (filters.favoriteOnly) {
    result = result.filter(function(food) {
      return isTrue(food.Favorite);
    });
  }

  /* งบประมาณ */
  if (filters.maxBudget !== '' && filters.maxBudget != null) {

    const budget = Number(filters.maxBudget);

    if (!isNaN(budget) && budget > 0) {
      result = result.filter(function(food) {
        const price = Number(food['ราคาโดยประมาณ']);

        return !isNaN(price) && price <= budget;
      });
    }
  }

  /* คะแนนขั้นต่ำ */
  if (filters.minRating !== '' && filters.minRating != null) {

    const minRating = Number(filters.minRating);

    if (!isNaN(minRating) && minRating > 0) {

      result = result.filter(function(food) {

        const foodRatings = ratings.filter(function(r) {
          return String(r.FoodID) === String(food.ID);
        });

        const scores = foodRatings
          .map(function(r) {
            return Number(r.Rating);
          })
          .filter(function(score) {
            return score >= 1 && score <= 5;
          });

        if (!scores.length) {
          return false;
        }

        const avg =
          scores.reduce(function(a, b) {
            return a + b;
          }, 0) / scores.length;

        return avg >= minRating;
      });
    }
  }

  /* ไม่เอาเมนูที่เพิ่งกิน */
  if (filters.excludeDays) {

    const days = Number(filters.excludeDays);

    if (!isNaN(days) && days > 0) {

      const cutoff = new Date();

      cutoff.setDate(cutoff.getDate() - days);

      const recentFoodIds = {};

      history.forEach(function(item) {

        const date = parseDate(item['วันที่กิน']);

        if (date && date >= cutoff) {
          recentFoodIds[String(item.FoodID)] = true;
        }
      });

      result = result.filter(function(food) {
        return !recentFoodIds[String(food.ID)];
      });
    }
  }

  /* สุ่ม */
  shuffleArray(result);

  return result;
}

/* =========================
   SMART DECISION
========================= */

function smartDecision(filters) {

  filters = filters || {};

  let candidates = randomFoods(filters);

  if (!candidates.length) {
    return null;
  }

  /*
   * ถ้ามีคะแนน ให้เพิ่มน้ำหนักเมนูคะแนนดี
   * แต่ยังคงมีความสุ่มอยู่
   */

  const ratings = getData('Ratings');

  const weighted = [];

  candidates.forEach(function(food) {

    const scores = ratings
      .filter(function(r) {
        return String(r.FoodID) === String(food.ID);
      })
      .map(function(r) {
        return Number(r.Rating);
      })
      .filter(function(score) {
        return score >= 1 && score <= 5;
      });

    let weight = 1;

    if (scores.length) {

      const avg =
        scores.reduce(function(a, b) {
          return a + b;
        }, 0) / scores.length;

      weight = Math.max(1, Math.round(avg * 2));
    }

    for (let i = 0; i < weight; i++) {
      weighted.push(food);
    }
  });

  shuffleArray(weighted);

  return weighted[0];
}

/* =========================
   RECOMMENDATION
========================= */

function getRecommendations() {

  const foods = getData('Foods')
    .filter(function(food) {
      return isTrue(food.Active);
    });

  const ratings = getData('Ratings');
  const history = getData('History');

  const now = new Date();

  /*
   * Smart Recommendation
   * ---------------------
   * ให้คะแนนแต่ละเมนูจาก:
   * 1) คะแนนเฉลี่ย
   * 2) จำนวนคน/จำนวนครั้งที่ให้คะแนน
   * 3) Favorite
   * 4) เมนูที่ไม่ได้กินมานาน จะได้คะแนนเพิ่ม
   * 5) เมนูที่เพิ่งกิน จะถูกลดคะแนน
   *
   * ยังคงมีความหลากหลาย เพราะมี random เล็กน้อย
   */

  const scored = foods.map(function(food) {

    const foodId = String(food.ID);

    const foodRatings = ratings
      .filter(function(r) {
        return String(r.FoodID) === foodId;
      })
      .map(function(r) {
        return Number(r.Rating);
      })
      .filter(function(score) {
        return score >= 1 && score <= 5;
      });

    const average = foodRatings.length
      ? foodRatings.reduce(function(a, b) {
          return a + b;
        }, 0) / foodRatings.length
      : 0;

    /* คะแนนพื้นฐานจาก Rating */
    let score = average > 0 ? average * 20 : 45;

    /* Favorite */
    if (isTrue(food.Favorite)) {
      score += 15;
    }

    /* มีคะแนนแล้ว เพิ่มความมั่นใจเล็กน้อย */
    if (foodRatings.length >= 2) {
      score += 5;
    }

    /* ดูประวัติการกินล่าสุด */
    const foodHistory = history
      .filter(function(item) {
        return String(item.FoodID) === foodId;
      })
      .map(function(item) {
        return parseDate(item['วันที่กิน']);
      })
      .filter(function(date) {
        return date;
      })
      .sort(function(a, b) {
        return b.getTime() - a.getTime();
      });

    if (foodHistory.length) {

      const lastEaten = foodHistory[0];
      const daysSince =
        (now.getTime() - lastEaten.getTime()) /
        (1000 * 60 * 60 * 24);

      /*
       * เพิ่งกิน -> ลดคะแนน
       * ไม่ได้กินนาน -> เพิ่มคะแนน
       */
      if (daysSince < 3) {
        score -= 35;
      } else if (daysSince < 7) {
        score -= 20;
      } else if (daysSince < 14) {
        score -= 5;
      } else if (daysSince < 30) {
        score += 8;
      } else {
        score += 15;
      }

    } else {
      /* เมนูที่ยังไม่เคยกิน ให้โอกาสค้นพบ */
      score += 10;
    }

    /* Random เล็กน้อย เพื่อไม่ให้ได้เมนูเดิมตลอด */
    score += Math.random() * 12;

    return {
      food: food,
      average: average,
      ratingCount: foodRatings.length,
      smartScore: score
    };
  });

  scored.sort(function(a, b) {
    return b.smartScore - a.smartScore;
  });

  /*
   * ส่ง 10 เมนูแรกกลับไปให้หน้า Home
   * หน้าเว็บสามารถสุ่ม/แสดงตามลำดับได้
   */
  return scored.slice(0, 10);
}

/* =========================
   STATISTICS
========================= */

function getStatistics() {

  const foods = getData('Foods');
  const ratings = getData('Ratings');
  const history = getData('History');

  const activeFoods = foods.filter(function(food) {
    return isTrue(food.Active);
  });

  const favoriteFoods = activeFoods.filter(function(food) {
    return isTrue(food.Favorite);
  });

  const scoreValues = ratings
    .map(function(item) {
      return Number(item.Rating);
    })
    .filter(function(score) {
      return score >= 1 && score <= 5;
    });

  const averageRating = scoreValues.length
    ? scoreValues.reduce(function(a, b) {
        return a + b;
      }, 0) / scoreValues.length
    : 0;

  const methodCount = {};

  activeFoods.forEach(function(food) {

    const method = String(food['วิธีทาน'] || 'ไม่ระบุ');

    methodCount[method] =
      (methodCount[method] || 0) + 1;
  });

  const typeCount = {};

  activeFoods.forEach(function(food) {

    const type = String(food['ประเภทอาหาร'] || 'ไม่ระบุ');

    typeCount[type] =
      (typeCount[type] || 0) + 1;
  });

  return {
    totalFoods: foods.length,
    activeFoods: activeFoods.length,
    favoriteFoods: favoriteFoods.length,
    totalRatings: ratings.length,
    totalHistory: history.length,
    averageRating: averageRating,
    methodCount: methodCount,
    typeCount: typeCount
  };
}

/* =========================
   UTILS
========================= */

function parseDate(value) {

  if (!value) {
    return null;
  }

  if (value instanceof Date) {
    return value;
  }

  const date = new Date(value);

  if (isNaN(date.getTime())) {
    return null;
  }

  return date;
}

function shuffleArray(array) {

  for (let i = array.length - 1; i > 0; i--) {

    const j = Math.floor(Math.random() * (i + 1));

    const temp = array[i];

    array[i] = array[j];
    array[j] = temp;
  }

  return array;
}
