Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        document.getElementById("sideload-msg").style.display = "none";
        document.getElementById("app-body").style.display = "flex";
        document.getElementById("run").onclick = run;
    }
});

async function run() {
    try {
        await Excel.run(async (context) => {
            // Lấy vùng dữ liệu đang bôi đen
            let range = context.workbook.getSelectedRange();
            range.load("values, rowCount");
            await context.sync();

            let values = range.values;
            let outputValues = [];

            // Duyệt qua từng dòng, dịch số và lưu vào mảng kết quả
            for (let i = 0; i < values.length; i++) {
                let cellValue = values[i][0];
                if (typeof cellValue === "number" && cellValue !== null) {
                    outputValues.push([docSoTienVND(cellValue)]);
                } else {
                    outputValues.push([""]); // Nếu không phải số thì để trống
                }
            }

            // Ghi kết quả ra cột ngay bên phải vùng bôi đen
            let targetRange = range.getOffsetRange(0, 1);
            targetRange.values = outputValues;
            
            await context.sync();
        });
    } catch (error) {
        console.error(error);
    }
}

// Logic chuyển số thành chữ tiếng Việt
function docSoTienVND(number) {
    if (number === 0) return "Không đồng";
    
    const units = ["không", "một", "hai", "ba", "bốn", "năm", "sáu", "bảy", "tám", "chín"];
    const blocks = ["", "nghìn", "triệu", "tỷ", "nghìn tỷ", "triệu tỷ"];
    
    let str = Math.abs(Math.round(number)).toString();
    let isNegative = number < 0;
    
    while (str.length % 3 !== 0) str = "0" + str;
    
    let words = [];
    let blockCount = str.length / 3;
    
    for (let i = 0; i < blockCount; i++) {
        let block = str.substring(i * 3, i * 3 + 3);
        let tram = parseInt(block[0]);
        let chuc = parseInt(block[1]);
        let donVi = parseInt(block[2]);
        
        if (tram === 0 && chuc === 0 && donVi === 0) continue;
        
        let blockWords = [];
        
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
        } else if (donVi > 0 || (chuc === 0 && tram === 0 && blockCount === 1)) {
            if (!(donVi === 1 && chuc > 1) && !(donVi === 5 && chuc > 0)) {
                 blockWords.push(units[donVi]);
            }
        }
        
        blockWords.push(blocks[blockCount - 1 - i]);
        words.push(blockWords.join(" "));
    }
    
    let result = words.join(" ").trim().replace(/\s+/g, ' '); 
    if (isNegative) result = "Âm " + result;
    
    return result.charAt(0).toUpperCase() + result.slice(1) + " đồng";
}