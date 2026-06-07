/* global Office, Word, Excel */
import lawData from '../../data.js'; 

let selectedLaws = []; // Lưu các văn bản đã tích chọn
let selectedTags = []; // Lưu các lĩnh vực đang chọn
let customLaws = JSON.parse(localStorage.getItem('customLaws')) || []; // Lấy dữ liệu người dùng đã thêm từ trước (nếu có)

// Đổi thành let để có thể cập nhật lại giá trị sau này
let CATEGORIES = [];

// Viết một hàm gộp chung dữ liệu gốc và dữ liệu người dùng tự thêm
function reloadCategoriesData() {
    const allLaws = [...lawData, ...customLaws];
    // Lọc lấy các lĩnh vực không trùng lặp và loại bỏ giá trị rỗng
    CATEGORIES = [...new Set(allLaws.map(item => item.category || item.linhvuc || item.LinhVuc))].filter(Boolean);
}

// ==========================================
// 1. KHỞI TẠO ADD-IN
// ==========================================
Office.onReady((info) => {
    try {
        if (info.host === Office.HostType.Word || info.host === Office.HostType.Excel) {

            reloadCategoriesData(); // Tải lại dữ liệu lĩnh vực mỗi khi khởi động
            initTagPicker();
            handleSearchAndFilter(); // Tải danh sách tìm kiếm

            // Gán sự kiện tìm kiếm
            const searchBox = document.getElementById("searchBox");
            if (searchBox) searchBox.oninput = handleSearchAndFilter;

            // Gán sự kiện nút bấm
            document.getElementById("insertBtn")?.addEventListener("click", insertToDocument);
            document.getElementById("clearBtn")?.addEventListener("click", resetAll);
        }
    } catch (err) {
        console.error("Lỗi khi khởi tạo Add-in:", err);
    }
});

// ==========================================
// 2. LOGIC TÌM KIẾM VÀ LỌC (SEARCH & FILTER)
// ==========================================
function handleSearchAndFilter() {
    const keyword = document.getElementById("searchBox")?.value.toLowerCase() || "";
    
    // Gộp dữ liệu gốc và dữ liệu người dùng tự thêm
    const allData = [...lawData, ...customLaws];
    
    const filtered = allData.filter(item => {
        // Tìm theo Trích yếu, Số ký hiệu, hoặc Loại văn bản
        const matchKeyword = (item.trichyeu || "").toLowerCase().includes(keyword) || 
                             (item.sokyhieu || "").toLowerCase().includes(keyword) ||
                             (item.coquanbanhanh || "").toLowerCase().includes(keyword) ||
                             (item.loaivanban || "").toLowerCase().includes(keyword);

        // Lọc theo Lĩnh vực (Hỗ trợ cả trường 'linhvuc', 'category', 'LinhVuc')
        const itemCategory = item.linhvuc || item.category || item.LinhVuc || "";
        const matchTag = selectedTags.length === 0 || selectedTags.includes(itemCategory);
        
        return matchKeyword && matchTag;
    });
    
    renderList(filtered);
}

function renderList(data) {
    const listDiv = document.getElementById("lawList");
    if (!listDiv) return;
    
    listDiv.innerHTML = ""; 
    
    if (data.length === 0) {
        listDiv.innerHTML = "<p style='padding:10px; color:gray;'>Không tìm thấy kết quả...</p>";
        return;
    }

    data.forEach((item) => {
        // 1. Tạo chuỗi hiển thị động từ các trường (dùng || để dự phòng nếu thiếu dữ liệu)
        const loai = item.loaivanban || "Văn bản";
        const so = item.sokyhieu || "[Chưa có số]";
        const coquanbanhanh = item.coquanbanhanh || "[Chưa có cơ quan]";
        const trich = item.trichyeu || "[Chưa có trích yếu]";
        // Lấy ngày ban hành và xử lý
        const ngay = item.ngaybanhanh || "";
        let chuoiNgayThang = "";
        if (ngay) {
            let parts = ngay.split('/'); // Tách chuỗi dd/mm/yyyy
            if (parts.length === 3) {
                let dayStr = parts[0];   // Giữ nguyên ngày (để lại số 0)
                let monthStr = parts[1]; // Xử lý tháng
                // Chỉ bỏ số 0 nếu tháng là từ 03 đến 09
                if (['03', '04', '05', '06', '07', '08', '09'].includes(monthStr)) {
                    monthStr = monthStr.substring(1); 
                }
                chuoiNgayThang = ` ngày ${dayStr} tháng ${monthStr} năm ${parts[2]}`;
            } else {
                chuoiNgayThang = ` ngày ${ngay}`; // Dự phòng
            }
        }
        let displayTitle = "";
        // Chuyển loại văn bản về chữ thường và xóa khoảng trắng thừa để so sánh cho chuẩn xác
        if (loai.toLowerCase().trim() === 'luật') {
            const trichYeuLuat = item.trichyeu ? item.trichyeu.trim() : "Luật"; 
            displayTitle = `${trichYeuLuat} số ${so}`;
        } else {
            displayTitle = `${loai} số ${so}${chuoiNgayThang} của ${coquanbanhanh} ${trich}`;
        }

        // Xóa khoảng trắng thừa ở 2 đầu (nếu có)
        displayTitle = displayTitle.trim();

        const itemDiv = document.createElement("div");
        
        // 2. Thay item.name thành displayTitle cho cả value, isChecked và hiển thị
        const isChecked = selectedLaws.includes(displayTitle) ? "checked" : "";

        itemDiv.innerHTML = `
            <label class="law-item-wrapper" style="display: flex; align-items: flex-start; padding: 8px; border-bottom: 1px solid #eee; cursor: pointer;">
                <input type="checkbox" class="law-item" value="${displayTitle}" ${isChecked} onchange="handleCheckboxChange(this)" style="margin-top: 4px; margin-right: 10px;">
                <span class="law-text" title="${displayTitle}" style="font-size: 14px; line-height: 1.4;">${displayTitle}</span>
            </label>
        `;
        listDiv.appendChild(itemDiv);
        
        // Gắn sự kiện click chuột phải (contextmenu) vào từng dòng
        const label = itemDiv.querySelector('.law-item-wrapper');
        label.addEventListener('contextmenu', function(e) {
          e.preventDefault(); 
          window.showContextMenu(e, item); 
        });    
    });
}

window.handleCheckboxChange = function(checkbox) {
    if (checkbox.checked) {
        if (!selectedLaws.includes(checkbox.value)) selectedLaws.push(checkbox.value);
    } else {
        selectedLaws = selectedLaws.filter(name => name !== checkbox.value);
    }
    // Cập nhật lại số lượng ngay khi tick/bỏ tick
    updateSelectionCount();
};

// ==========================================
// 3. LOGIC CHO MENU LĨNH VỰC (TAG PICKER)
// ==========================================
function initTagPicker() {
    const tagPicker = document.getElementById("tagPicker");
    const tagDropdown = document.getElementById("tagDropdown");
    const tagDropdownList = document.getElementById("tagDropdownList");

    CATEGORIES.forEach(cat => {
        const item = document.createElement("div");
        item.className = "dropdown-item";
        item.innerHTML = `<span>${cat}</span> <span class="check-mark">✔</span>`;
        item.onclick = (e) => {
            e.stopPropagation(); 
            toggleTag(cat, item);
        };
        tagDropdownList.appendChild(item);
    });

    tagPicker.onclick = (e) => {
        e.stopPropagation();
        tagDropdown.style.display = tagDropdown.style.display === "block" ? "none" : "block";
    };

    document.addEventListener("click", (e) => {
        if (!tagPicker.contains(e.target) && !tagDropdown.contains(e.target)) {
            tagDropdown.style.display = "none";
        }
    });
}

function toggleTag(category, element) {
    const index = selectedTags.indexOf(category);
    if (index > -1) {
        selectedTags.splice(index, 1);
        element.classList.remove("selected");
    } else {
        selectedTags.push(category);
        element.classList.add("selected");
    }
    renderSelectedTags();
    handleSearchAndFilter(); 
}

function renderSelectedTags() {
    const container = document.getElementById("selectedTags");
    const input = document.getElementById("tagInput");
    
    container.innerHTML = ""; 
    
    if (selectedTags.length > 0) {
        input.style.display = "none"; 
        container.style.display = "flex"; 

        let visibleCount = 0;
        const moreTag = document.createElement("div");
        moreTag.className = "tag tag-more";
        
        for (let i = 0; i < selectedTags.length; i++) {
            const tag = selectedTags[i];
            const tagEl = document.createElement("div");
            tagEl.className = "tag";
            tagEl.innerHTML = `${tag} <span style="margin-left:5px; cursor:pointer;" onclick="removeTag(event, '${tag}')">✖</span>`;
            
            container.appendChild(tagEl);
            
            // KIỂM TRA TRÀN KHUNG (OVERFLOW)
            if (container.scrollWidth > container.clientWidth) {
                container.removeChild(tagEl);
                moreTag.innerHTML = `+${selectedTags.length - visibleCount}`;
                container.appendChild(moreTag);
                
                while (container.scrollWidth > container.clientWidth && visibleCount > 0) {
                    container.removeChild(container.childNodes[visibleCount - 1]);
                    visibleCount--;
                    moreTag.innerHTML = `+${selectedTags.length - visibleCount}`;
                }
                break; 
            }
            visibleCount++; 
        }
    } else {
        input.style.display = "block"; 
        container.style.display = "none";
    }
}

window.removeTag = function(event, tag) {
    event.stopPropagation();
    const items = document.querySelectorAll(".dropdown-item");
    items.forEach(item => {
        if (item.querySelector("span").innerText === tag) toggleTag(tag, item);
    });
};

// ==========================================
// 4. CÁC NÚT CHỨC NĂNG (CHÈN / RESET)
// ==========================================
function resetAll() {
    selectedLaws = []; 
    handleSearchAndFilter(); 
    updateSelectionCount();
}

async function insertToDocument() {
  
  const addPrefix = document.getElementById("prefixCheckbox")?.checked;
  const addSuffix = document.getElementById("suffixCheckbox")?.checked;

  const processedLaws = selectedLaws.map(law => {
    let text = law;
    if (addPrefix) text = "Căn cứ " + text;
    if (addSuffix) text = text + ";";
    return text;
  });

  const textToInsert = processedLaws.join("\n");

  try {
    if (Office.context.host === Office.HostType.Word) {
      await Word.run(async (context) => {
        const range = context.document.getSelection();
        range.insertText(textToInsert + "\n", Word.InsertLocation.replace);
        await context.sync();
      });
    } else if (Office.context.host === Office.HostType.Excel) {
      await Excel.run(async (context) => {
        const range = context.workbook.getSelectedRange();
        const values = processedLaws.map(line => [line]);
        range.getResizedRange(values.length - 1, 0).values = values;
        await context.sync();
      });
    }

    resetAll();
    
  } catch (error) {
    console.error("Lỗi chèn văn bản:", error);
  }
}

// ==========================================
// 5. CHUYỂN TAB
// ==========================================
window.openTab = function(evt, tabName) {
    const tabContents = document.getElementsByClassName("tab-content");
    for (let i = 0; i < tabContents.length; i++) tabContents[i].classList.remove("active");

    const tabLinks = document.getElementsByClassName("tab-link");
    for (let i = 0; i < tabLinks.length; i++) tabLinks[i].classList.remove("active");

    document.getElementById(tabName).classList.add("active");
    evt.currentTarget.classList.add("active");
}

// ==========================================
// 6. CẬP NHẬT THỐNG KÊ
// ==========================================
function updateSelectionCount() {
    const countEl = document.getElementById("selectionCount");
    if (countEl) {
        countEl.innerText = `Đã chọn: ${selectedLaws.length} văn bản`;
    }
}

// ==========================================
// 7. MENU CHUỘT PHẢI & MỞ DIALOG
// ==========================================
let currentRightClickedItem = null;
let dialog; // Biến lưu trữ cửa sổ dialog

window.showContextMenu = function(e, item) {
    currentRightClickedItem = item;
    const menu = document.getElementById('contextMenu');
    menu.style.display = "block";
    const menuWidth = menu.offsetWidth;
    const menuHeight = menu.offsetHeight;
    const windowWidth = window.innerWidth;
    const windowHeight = window.innerHeight;
    if (windowWidth - e.clientX < menuWidth) {
        menu.style.left = (e.clientX - menuWidth) + "px";
    } else {
        menu.style.left = e.clientX + "px";
    }
    if (windowHeight - e.clientY < menuHeight) {
        menu.style.top = (e.clientY - menuHeight) + "px";
    } else {
        menu.style.top = e.clientY + "px";
    }
};

document.addEventListener('click', function(e) {
    const menu = document.getElementById('contextMenu');
    if (menu) menu.style.display = 'none';
});

window.openModal = function(type) {
    if (!currentRightClickedItem) return;

    // 1. Đóng gói dữ liệu để gửi sang cửa sổ dialog
    const dataToSend = {
        type: type, 
        title: type === 'details' ? 'Chi tiết văn bản' : (type === 'expired' ? 'Báo văn bản hết hạn' : 'Báo lỗi văn bản'),
        item: currentRightClickedItem
    };
    
    // 2. Mã hóa dữ liệu để gắn vào URL 
    const encodedData = encodeURIComponent(JSON.stringify(dataToSend));
    const dialogUrl = window.location.origin + '/dialog.html?data=' + encodedData;

    // 3. Mở cửa sổ nổi lên giữa màn hình
    Office.context.ui.displayDialogAsync(
        dialogUrl, 
        { height: 50, width: 35, displayInIframe: true }, 
        function (asyncResult) {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                console.error("Lỗi khi mở cửa sổ:", asyncResult.error.message);
            } else {
                dialog = asyncResult.value;
                
                dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
                    if (arg.message === "close") {
                        dialog.close();
                    }
                });
            }
        }
    );
    
    // Ẩn menu chuột phải đi sau khi click
    document.getElementById('contextMenu').style.display = 'none';
};

// ==========================================
// 8. THÊM, QUẢN LÝ VĂN BẢN
// ==========================================

// Sự kiện nút Thêm văn bản mới
document.getElementById("addLawBtn").onclick = function() {
    const dialogUrl = window.location.origin + '/add-law-dialog.html';
    
    Office.context.ui.displayDialogAsync(dialogUrl, { height: 72, width: 60 }, function(asyncResult) {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) return;

        const dialogObj = asyncResult.value;
        
        // Lắng nghe tin nhắn từ cửa sổ Dialog gửi về
        dialogObj.addEventHandler(Office.EventType.DialogMessageReceived, function(arg) {
            // Trường hợp 1: Nhấn Hủy bỏ (gửi chuỗi "close")
            if (arg.message === "close") {
                dialogObj.close();
                return; // Kết thúc luôn, không chạy đoạn code dưới
            }

            // Trường hợp 2: Nhấn Lưu (gửi chuỗi JSON dữ liệu)
            try {
                const result = JSON.parse(arg.message);
                if (result.action === 'addLaw') {
                    // 1. Lưu dữ liệu mới
                    customLaws.push(result.data);
                    localStorage.setItem('customLaws', JSON.stringify(customLaws));
                    dialogObj.close();
                    // 2. Reload cửa sổ
                    window.location.reload();
                }
         
            } catch (e) {
                // Nếu tin nhắn không phải JSON thì có thể là lệnh đóng khác
                dialogObj.close();
            }
        });
    });
};

// Sự kiện nút Quản lý văn bản
document.getElementById("manageLawBtn").onclick = function() {
    const dialogUrl = window.location.origin + '/manage-laws.html';
    
    Office.context.ui.displayDialogAsync(dialogUrl, { height: 72, width: 60 }, function(asyncResult) {
        const dialogObj = asyncResult.value;
        dialogObj.addEventHandler(Office.EventType.DialogMessageReceived, function(arg) {
            const result = JSON.parse(arg.message);
            
            // 1. Khi cửa sổ Quản lý mở lên, gửi dữ liệu sang cho nó
            if (result.action === 'ready') {
                dialogObj.messageChild(JSON.stringify({ 
                    action: 'loadData', 
                    data: customLaws
                }));
            }
            // 2. KHI CÓ THAO TÁC XÓA/SỬA TỪ CỬA SỔ QUẢN LÝ GỬI VỀ
            else if (result.action === 'refresh') {
                // 1. Ghi đè danh sách hiện tại bằng danh sách mới (đã mất phần tử bị xóa)
                customLaws = result.data; 
                // 2. Taskpane tự lưu danh sách mới này vào bộ nhớ
                localStorage.setItem('customLaws', JSON.stringify(customLaws));
                // 3. Reload cửa sổ
                window.location.reload();
            }
            else if (result.action === 'close') {
                dialogObj.close();
            }
        });
    });
};