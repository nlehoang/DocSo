/**
 * Chuyển đổi số thành chữ tiếng Việt (xử lý phần nguyên)
 * @param {number} number - Số cần chuyển đổi
 * @returns {string} - Chuỗi chữ tiếng Việt
 */
function convertIntegerPart(number) {
    if (number === 0) return "không";

    var units = ["không", "một", "hai", "ba", "bốn", "năm", "sáu", "bảy", "tám", "chín"];
    var blocks = ["", "nghìn", "triệu", "tỷ", "nghìn tỷ", "triệu tỷ"];

    var str = Math.abs(Math.round(number)).toString();
    var isNegative = number < 0;

    while (str.length % 3 !== 0) str = "0" + str;

    var words = [];
    var blockCount = str.length / 3;

    for (var i = 0; i < blockCount; i++) {
        var block = str.substring(i * 3, i * 3 + 3);
        var tram = parseInt(block[0], 10);
        var chuc = parseInt(block[1], 10);
        var donVi = parseInt(block[2], 10);

        if (tram === 0 && chuc === 0 && donVi === 0) continue;

        var blockWords = [];

        if (tram > 0 || (blockCount > 1 && i > 0)) {
            blockWords.push(units[tram] + " trăm");
        }

        if (chuc === 0 && donVi > 0 && (tram > 0 || (blockCount > 1 && i > 0))) {
            blockWords.push("lẻ");
        } else if (chuc === 1) {
            blockWords.push("mười");
        } else if (chuc > 1) {
            blockWords.push(units[chuc] + " mươi");
        }

        if (donVi === 1 && chuc > 1) {
            blockWords.push("mốt");
        } else if (donVi === 5 && chuc > 0) {
            blockWords.push("lăm");
        } else if (donVi === 4 && chuc > 1) {
            blockWords.push("tư");
        } else if (donVi > 0) {
            blockWords.push(units[donVi]);
        }

        var blockIndex = blockCount - 1 - i;
        if (blockIndex > 0 && blockIndex < blocks.length) {
            blockWords.push(blocks[blockIndex]);
        }
        words.push(blockWords.join(" "));
    }

    var result = words.join(" ").trim().replace(/\s+/g, " ");
    if (isNegative) result = "Âm " + result;

    return result;
}

/**
 * Chuyển đổi số tiền thành chữ tiếng Việt (hỗ trợ số thập phân)
 * @customfunction VND
 * @param {number} number Số tiền cần chuyển đổi
 * @returns {string} Chuỗi văn bản tương ứng
 */
function docSoTienVND(number) {
    try {
        if (typeof number !== "number" || isNaN(number)) {
            return "Lỗi: Giá trị không hợp lệ";
        }

        if (number === 0) return "Không đồng";

        var isNegative = number < 0;
        var absNumber = Math.abs(number);

        // Tách phần nguyên và phần thập phân
        var integerPart = Math.floor(absNumber);
        var decimalPart = Math.round((absNumber - integerPart) * 100); // Lấy 2 chữ số thập phân

        var result = "";

        // Xử lý phần nguyên
        if (integerPart > 0) {
            result = convertIntegerPart(integerPart);
            result = result.charAt(0).toUpperCase() + result.slice(1);
        } else {
            result = "Không";
        }

        // Xử lý phần thập phân
        if (decimalPart > 0) {
            var decimalWords = convertIntegerPart(decimalPart);
            result += " phẩy " + decimalWords;
        }

        result += " đồng";

        if (isNegative) {
            result = "Âm " + result;
        }

        return result;
    } catch (error) {
        return "Lỗi: " + error.message;
    }
}

/**
 * Chuyển đổi số tiền thành chữ tiếng Việt - Đô la Mỹ (USD)
 * @customfunction USD
 * @param {number} number Số tiền cần chuyển đổi
 * @returns {string} Chuỗi văn bản tương ứng
 */
function docSoUSD(number) {
    try {
        if (typeof number !== "number" || isNaN(number)) {
            return "Lỗi: Giá trị không hợp lệ";
        }

        if (number === 0) return "Không đô la Mỹ";

        var isNegative = number < 0;
        var absNumber = Math.abs(number);

        var integerPart = Math.floor(absNumber);
        var decimalPart = Math.round((absNumber - integerPart) * 100);

        var result = "";

        if (integerPart > 0) {
            result = convertIntegerPart(integerPart);
            result = result.charAt(0).toUpperCase() + result.slice(1);
        } else {
            result = "Không";
        }

        if (decimalPart > 0) {
            var decimalWords = convertIntegerPart(decimalPart);
            result += " phẩy " + decimalWords;
        }

        result += " đô la Mỹ";

        if (isNegative) {
            result = "Âm " + result;
        }

        return result;
    } catch (error) {
        return "Lỗi: " + error.message;
    }
}

/**
 * Chuyển đổi số tiền thành chữ tiếng Việt - Đô la Canada (CAD)
 * @customfunction CAD
 * @param {number} number Số tiền cần chuyển đổi
 * @returns {string} Chuỗi văn bản tương ứng
 */
function docSoCAD(number) {
    try {
        if (typeof number !== "number" || isNaN(number)) {
            return "Lỗi: Giá trị không hợp lệ";
        }

        if (number === 0) return "Không đô la Canada";

        var isNegative = number < 0;
        var absNumber = Math.abs(number);

        var integerPart = Math.floor(absNumber);
        var decimalPart = Math.round((absNumber - integerPart) * 100);

        var result = "";

        if (integerPart > 0) {
            result = convertIntegerPart(integerPart);
            result = result.charAt(0).toUpperCase() + result.slice(1);
        } else {
            result = "Không";
        }

        if (decimalPart > 0) {
            var decimalWords = convertIntegerPart(decimalPart);
            result += " phẩy " + decimalWords;
        }

        result += " đô la Canada";

        if (isNegative) {
            result = "Âm " + result;
        }

        return result;
    } catch (error) {
        return "Lỗi: " + error.message;
    }
}

/**
 * Chuyển đổi số tiền thành chữ tiếng Việt - Nhân dân tệ (CNY)
 * @customfunction CNY
 * @param {number} number Số tiền cần chuyển đổi
 * @returns {string} Chuỗi văn bản tương ứng
 */
function docSoCNY(number) {
    try {
        if (typeof number !== "number" || isNaN(number)) {
            return "Lỗi: Giá trị không hợp lệ";
        }

        if (number === 0) return "Không nhân dân tệ";

        var isNegative = number < 0;
        var absNumber = Math.abs(number);

        var integerPart = Math.floor(absNumber);
        var decimalPart = Math.round((absNumber - integerPart) * 100);

        var result = "";

        if (integerPart > 0) {
            result = convertIntegerPart(integerPart);
            result = result.charAt(0).toUpperCase() + result.slice(1);
        } else {
            result = "Không";
        }

        if (decimalPart > 0) {
            var decimalWords = convertIntegerPart(decimalPart);
            result += " phẩy " + decimalWords;
        }

        result += " nhân dân tệ";

        if (isNegative) {
            result = "Âm " + result;
        }

        return result;
    } catch (error) {
        return "Lỗi: " + error.message;
    }
}

/**
 * Chuyển đổi số tiền thành chữ tiếng Việt - Yên Nhật (JPY)
 * @customfunction JPY
 * @param {number} number Số tiền cần chuyển đổi
 * @returns {string} Chuỗi văn bản tương ứng
 */
function docSoJPY(number) {
    try {
        if (typeof number !== "number" || isNaN(number)) {
            return "Lỗi: Giá trị không hợp lệ";
        }

        if (number === 0) return "Không yên Nhật";

        var isNegative = number < 0;
        var absNumber = Math.abs(number);

        var integerPart = Math.floor(absNumber);
        var decimalPart = Math.round((absNumber - integerPart) * 100);

        var result = "";

        if (integerPart > 0) {
            result = convertIntegerPart(integerPart);
            result = result.charAt(0).toUpperCase() + result.slice(1);
        } else {
            result = "Không";
        }

        if (decimalPart > 0) {
            var decimalWords = convertIntegerPart(decimalPart);
            result += " phẩy " + decimalWords;
        }

        result += " yên Nhật";

        if (isNegative) {
            result = "Âm " + result;
        }

        return result;
    } catch (error) {
        return "Lỗi: " + error.message;
    }
}

/**
 * Chuyển đổi số tiền thành chữ tiếng Việt - Won Hàn Quốc (KRW)
 * @customfunction KRW
 * @param {number} number Số tiền cần chuyển đổi
 * @returns {string} Chuỗi văn bản tương ứng
 */
function docSoKRW(number) {
    try {
        if (typeof number !== "number" || isNaN(number)) {
            return "Lỗi: Giá trị không hợp lệ";
        }

        if (number === 0) return "Không won Hàn Quốc";

        var isNegative = number < 0;
        var absNumber = Math.abs(number);

        var integerPart = Math.floor(absNumber);
        var decimalPart = Math.round((absNumber - integerPart) * 100);

        var result = "";

        if (integerPart > 0) {
            result = convertIntegerPart(integerPart);
            result = result.charAt(0).toUpperCase() + result.slice(1);
        } else {
            result = "Không";
        }

        if (decimalPart > 0) {
            var decimalWords = convertIntegerPart(decimalPart);
            result += " phẩy " + decimalWords;
        }

        result += " won Hàn Quốc";

        if (isNegative) {
            result = "Âm " + result;
        }

        return result;
    } catch (error) {
        return "Lỗi: " + error.message;
    }
}

/**
 * Chuyển đổi số tiền thành chữ tiếng Việt - Bạt Thái Lan (THB)
 * @customfunction THB
 * @param {number} number Số tiền cần chuyển đổi
 * @returns {string} Chuỗi văn bản tương ứng
 */
function docSoTHB(number) {
    try {
        if (typeof number !== "number" || isNaN(number)) {
            return "Lỗi: Giá trị không hợp lệ";
        }

        if (number === 0) return "Không bạt Thái Lan";

        var isNegative = number < 0;
        var absNumber = Math.abs(number);

        var integerPart = Math.floor(absNumber);
        var decimalPart = Math.round((absNumber - integerPart) * 100);

        var result = "";

        if (integerPart > 0) {
            result = convertIntegerPart(integerPart);
            result = result.charAt(0).toUpperCase() + result.slice(1);
        } else {
            result = "Không";
        }

        if (decimalPart > 0) {
            var decimalWords = convertIntegerPart(decimalPart);
            result += " phẩy " + decimalWords;
        }

        result += " bạt Thái Lan";

        if (isNegative) {
            result = "Âm " + result;
        }

        return result;
    } catch (error) {
        return "Lỗi: " + error.message;
    }
}

/**
 * Chuyển đổi số tiền thành chữ tiếng Việt - Đô la Singapore (SGD)
 * @customfunction SGD
 * @param {number} number Số tiền cần chuyển đổi
 * @returns {string} Chuỗi văn bản tương ứng
 */
function docSoSGD(number) {
    try {
        if (typeof number !== "number" || isNaN(number)) {
            return "Lỗi: Giá trị không hợp lệ";
        }

        if (number === 0) return "Không đô la Singapore";

        var isNegative = number < 0;
        var absNumber = Math.abs(number);

        var integerPart = Math.floor(absNumber);
        var decimalPart = Math.round((absNumber - integerPart) * 100);

        var result = "";

        if (integerPart > 0) {
            result = convertIntegerPart(integerPart);
            result = result.charAt(0).toUpperCase() + result.slice(1);
        } else {
            result = "Không";
        }

        if (decimalPart > 0) {
            var decimalWords = convertIntegerPart(decimalPart);
            result += " phẩy " + decimalWords;
        }

        result += " đô la Singapore";

        if (isNegative) {
            result = "Âm " + result;
        }

        return result;
    } catch (error) {
        return "Lỗi: " + error.message;
    }
}

// Export các function cho Excel Custom Functions
if (typeof CustomFunctions !== "undefined") {
    CustomFunctions.associate("VND", docSoTienVND);
    CustomFunctions.associate("USD", docSoUSD);
    CustomFunctions.associate("CAD", docSoCAD);
    CustomFunctions.associate("CNY", docSoCNY);
    CustomFunctions.associate("JPY", docSoJPY);
    CustomFunctions.associate("KRW", docSoKRW);
    CustomFunctions.associate("THB", docSoTHB);
    CustomFunctions.associate("SGD", docSoSGD);
}

// Export global scope
if (typeof window !== "undefined") {
    window.docSoTienVND = docSoTienVND;
    window.docSoUSD = docSoUSD;
    window.docSoCAD = docSoCAD;
    window.docSoCNY = docSoCNY;
    window.docSoJPY = docSoJPY;
    window.docSoKRW = docSoKRW;
    window.docSoTHB = docSoTHB;
    window.docSoSGD = docSoSGD;
}
