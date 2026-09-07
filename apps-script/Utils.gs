/* -------------------- Утилиты -------------------- */

function columnToLetter_(column) {
  let result = '';
  let n = Number(column);

  while (n > 0) {
    const rem = (n - 1) % 26;
    result =
      String.fromCharCode(65 + rem) +
      result;
    n = Math.floor((n - 1) / 26);
  }

  return result;
}

