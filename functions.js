/**
 * Chuyển đổi số tiền thành chữ tiếng Việt
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

        return result.charAt(0).toUpperCase() + result.slice(1) + " đồng";
    } catch (error) {
        return "Lỗi: " + error.message;
    }
}

/**
 * Test function to check if runtime is working
 * @customfunction PING
 * @returns {string}
 */
function ping() {
    return "PONG";
}

// Export các function cho Excel Custom Functions
if (typeof CustomFunctions !== "undefined") {
    CustomFunctions.associate("VND", docSoTienVND);
    CustomFunctions.associate("PING", ping);
}

// Export global scope
if (typeof window !== "undefined") {
    window.docSoTienVND = docSoTienVND;
    window.ping = ping;
}
