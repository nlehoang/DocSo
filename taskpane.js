Office.onReady(function (info) {
    if (info.host === Office.HostType.Excel) {
        document.getElementById("sideload-msg").style.display = "none";
        document.getElementById("app-body").style.display = "flex";
        document.getElementById("run").onclick = run;
    }
});

async function run() {
    var status = document.getElementById("status-msg");
    status.innerText = "";
    status.className = "ms-font-s";

    try {
        await Excel.run(async function (context) {
            // Lấy vùng dữ liệu đang bôi đen
            var range = context.workbook.getSelectedRange();
            range.load("values,columnCount");
            await context.sync();

            var values = range.values;
            var colCount = range.columnCount;
            var outputValues = [];

            // Duyệt qua từng dòng
            for (var i = 0; i < values.length; i++) {
                var row = [];
                for (var j = 0; j < colCount; j++) {
                    var cellValue = values[i][j];
                    if (typeof cellValue === "number") {
                        row.push(docSoTienVND(cellValue));
                    } else {
                        row.push(""); // Nếu không phải số thì để trống
                    }
                }
                outputValues.push(row);
            }

            // Ghi kết quả ra vùng ngay bên phải vùng bôi đen
            var targetRange = range.getOffsetRange(0, colCount);
            targetRange.values = outputValues;

            status.className = "ms-font-s status-success";
            status.innerText = "Đã chuyển đổi thành công!";
            await context.sync();
        });
    } catch (error) {
        status.className = "ms-font-s status-error";
        status.innerText = "Lỗi: " + error.message;
        console.error(error);
    }
}

/**
 * Logic chuyển đổi số thành chữ tiếng Việt
 * @param {number} number - Con số cần chuyển đổi
 * @returns {string} - Chuỗi văn bản tương ứng (ví dụ: "Một trăm nghìn đồng")
 */
function docSoTienVND(number) {
    if (number === 0) return "Không đồng";

    var units = ["không", "một", "hai", "ba", "bốn", "năm", "sáu", "bảy", "tám", "chín"];
    var blocks = ["", "nghìn", "triệu", "tỷ", "nghìn tỷ", "triệu tỷ"];

    // Làm tròn và lấy giá trị tuyệt đối
    var str = Math.abs(Math.round(number)).toString();
    var isNegative = number < 0;

    // Thêm số 0 vào đầu để đủ các nhóm 3 chữ số
    while (str.length % 3 !== 0) str = "0" + str;

    var words = [];
    var blockCount = str.length / 3;

    for (var i = 0; i < blockCount; i++) {
        var block = str.substring(i * 3, i * 3 + 3);
        var tram = parseInt(block[0], 10);
        var chuc = parseInt(block[1], 10);
        var donVi = parseInt(block[2], 10);

        // Nếu cả nhóm là 000 thì bỏ qua
        if (tram === 0 && chuc === 0 && donVi === 0) continue;

        var blockWords = [];

        // Xử lý hàng trăm
        if (tram > 0 || (blockCount > 1 && i > 0)) {
            blockWords.push(units[tram] + " trăm");
        }

        // Xử lý hàng chục
        if (chuc === 0 && donVi > 0 && (tram > 0 || (blockCount > 1 && i > 0))) {
            blockWords.push("lẻ");
        } else if (chuc === 1) {
            blockWords.push("mười");
        } else if (chuc > 1) {
            blockWords.push(units[chuc] + " mươi");
        }

        // Xử lý hàng đơn vị (các trường hợp đặc biệt: mốt, lăm, tư)
        if (donVi === 1 && chuc > 1) {
            blockWords.push("mốt");
        } else if (donVi === 5 && chuc > 0) {
            blockWords.push("lăm");
        } else if (donVi === 4 && chuc > 1) {
            blockWords.push("tư");
        } else if (donVi > 0) {
            blockWords.push(units[donVi]);
        }

        // Thêm đơn vị khối (nghìn, triệu, tỷ...)
        var blockIndex = blockCount - 1 - i;
        if (blockIndex > 0 && blockIndex < blocks.length) {
            blockWords.push(blocks[blockIndex]);
        }
        words.push(blockWords.join(" "));
    }

    // Kết hợp các nhóm lại
    var result = words.join(" ").trim().replace(/\s+/g, " ");
    if (isNegative) result = "Âm " + result;

    // Viết hoa chữ cái đầu và thêm "đồng"
    return result.charAt(0).toUpperCase() + result.slice(1) + " đồng";
}