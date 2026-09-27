// --- QUẢN LÝ TRẠNG THÁI ỨNG DỤNG ---
let topicsData = {}; // Cấu trúc: { "Unit 1.xlsx": [{word: "apple", meaning: "quả táo"}, ...] }
let currentStudyWords = [];
let currentMode = 'flashcard';

// Biến cho Flashcard
let fcIndex = 0;

// Biến cho Quiz
let quizIndex = 0;
let quizScore = 0;
let currentQuizQuestion = null;

// --- CÁC DOM ELEMENTS ---
const fileInput = document.getElementById('file-input');
const btnAddFile = document.getElementById('btn-add-file');
const btnDownloadSample = document.getElementById('btn-download-sample');
const topicsContainerCard = document.getElementById('topics-container-card');
const topicsListEl = document.getElementById('topics-list');
const statsOverview = document.getElementById('stats-overview');
const checkboxSelectAll = document.getElementById('checkbox-select-all');
const btnDeleteAll = document.getElementById('btn-delete-all');
const btnStartStudy = document.getElementById('btn-start-study');

const studySection = document.getElementById('study-section');
const btnBackToHome = document.getElementById('btn-back-to-home');
const studyTitle = document.getElementById('study-title');

const modeFlashcardContent = document.getElementById('mode-flashcard-content');
const modeQuizContent = document.getElementById('mode-quiz-content');
const studyCompleteContent = document.getElementById('study-complete-content');

// Modal
const confirmModal = document.getElementById('confirm-modal');
const modalBtnCancel = document.getElementById('modal-btn-cancel');
const modalBtnConfirm = document.getElementById('modal-btn-confirm');

// --- 1. TẢI FILE XLSX MẪU ---
btnDownloadSample.addEventListener('click', () => {
    const sampleData = [
        { "Từ": "apple", "Nghĩa": "quả táo" },
        { "Từ": "book", "Nghĩa": "quyển sách" },
        { "Từ": "teacher", "Nghĩa": "giáo viên" },
        { "Từ": "school", "Nghĩa": "trường học" }
    ];

    const worksheet = XLSX.utils.json_to_sheet(sampleData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "TuVung");
    XLSX.writeFile(workbook, "mau_tu_vung.xlsx");
});

// --- 2. UPLOAD & ĐỌC NHIỀU FILE XLSX ---
btnAddFile.addEventListener('click', () => {
    fileInput.click();
});

fileInput.addEventListener('change', async (e) => {
    const files = e.target.files;
    if (!files || files.length === 0) return;

    for (let file of files) {
        try {
            const data = await readExcelFile(file);
            if (data && data.length > 0) {
                topicsData[file.name] = data;
            }
        } catch (err) {
            console.error("Lỗi khi đọc file:", file.name, err);
            alert(`Không thể đọc file "${file.name}". Hãy chắc chắn đó là file Excel chuẩn!`);
        }
    }

    // Reset input để có thể chọn lại chính file đó nếu cần
    fileInput.value = '';
    updateUI();
});

function readExcelFile(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                const firstSheetName = workbook.SheetNames[0];
                const worksheet = workbook.Sheets[firstSheetName];
                
                // Chuyển sheet thành JSON dạng mảng các object
                const jsonRows = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
                
                if (jsonRows.length < 2) {
                    resolve([]);
                    return;
                }

                // Dòng đầu tiên là tiêu đề (Cột 1: Từ, Cột 2: Nghĩa)
                let wordList = [];
                for (let i = 1; i < jsonRows.length; i++) {
                    const row = jsonRows[i];
                    if (row && row[0] !== undefined && row[1] !== undefined) {
                        const word = String(row[0]).trim();
                        const meaning = String(row[1]).trim();
                        if (word && meaning) {
                            wordList.push({ word, meaning });
                        }
                    }
                }
                resolve(wordList);
            } catch (error) {
                reject(error);
            }
        };
        reader.onerror = (error) => reject(error);
        reader.readAsArrayBuffer(file);
    });
}

// --- 3. CẬP NHẬT GIAO DIỆN CHỦ ĐỀ & THỐNG KÊ ---
function updateUI() {
    const topicNames = Object.keys(topicsData);
    const totalTopics = topicNames.length;
    
    let totalWords = 0;
    topicNames.forEach(name => {
        totalWords += topicsData[name].length;
    });

    // Cập nhật thanh tổng quan
    statsOverview.textContent = `Đã nạp: ${totalTopics} bộ từ vựng — ${totalWords} từ`;

    if (totalTopics === 0) {
        topicsContainerCard.style.display = 'none';
        return;
    }

    topicsContainerCard.style.display = 'block';
    
    // Render danh sách chủ đề
    topicsListEl.innerHTML = '';
    topicNames.forEach(name => {
        const wordCount = topicsData[name].length;
        const topicItem = document.createElement('div');
        topicItem.className = 'topic-item';
        topicItem.innerHTML = `
            <label class="topic-label">
                <input type="checkbox" class="topic-checkbox" value="${name}" checked>
                <span>📁 ${name}</span> 
                <small style="color: #64748b; font-weight: normal;">(${wordCount} từ)</small>
            </label>
            <button class="btn btn-danger-text btn-delete-topic" data-name="${name}" title="Xóa chủ đề">🗑️</button>
        `;
        topicsListEl.appendChild(topicItem);
    });

    // Gắn sự kiện xóa từng chủ đề
    document.querySelectorAll('.btn-delete-topic').forEach(btn => {
        btn.addEventListener('click', (e) => {
            const topicName = e.currentTarget.getAttribute('data-name');
            delete topicsData[topicName];
            updateUI();
        });
    });

    // Cập nhật trạng thái checkbox "Chọn tất cả"
    updateSelectAllCheckboxState();
}

// Xử lý chọn tất cả checkbox
checkboxSelectAll.addEventListener('change', (e) => {
    const isChecked = e.target.checked;
    document.querySelectorAll('.topic-checkbox').forEach(cb => {
        cb.checked = isChecked;
    });
});

topicsListEl.addEventListener('change', () => {
    updateSelectAllCheckboxState();
});

function updateSelectAllCheckboxState() {
    const checkboxes = document.querySelectorAll('.topic-checkbox');
    if (checkboxes.length === 0) return;
    const allChecked = Array.from(checkboxes).every(cb => cb.checked);
    checkboxSelectAll.checked = allChecked;
}

// --- 4. XÓA TẤT CẢ CÓ XÁC NHẬN ---
btnDeleteAll.addEventListener('click', () => {
    confirmModal.style.display = 'flex';
});

modalBtnCancel.addEventListener('click', () => {
    confirmModal.style.display = 'none';
});

modalBtnConfirm.addEventListener('click', () => {
    topicsData = {};
    confirmModal.style.display = 'none';
    updateUI();
});

// --- 5. BẮT ĐẦU HỌC ---
btnStartStudy.addEventListener('click', () => {
    // Thu thập các chủ đề được chọn
    const selectedCheckboxes = document.querySelectorAll('.topic-checkbox:checked');
    if (selectedCheckboxes.length === 0) {
        alert('Vui lòng chọn ít nhất một bộ từ vựng để học nhé! 😊');
        return;
    }

    currentStudyWords = [];
    selectedCheckboxes.forEach(cb => {
        const topicName = cb.value;
        if (topicsData[topicName]) {
            currentStudyWords = currentStudyWords.concat(topicsData[topicName]);
        }
    });

    // Trộn ngẫu nhiên danh sách từ vựng khi học
    shuffleArray(currentStudyWords);

    // Lấy chế độ học
    const selectedModeRadio = document.querySelector('input[name="study-mode"]:checked');
    currentMode = selectedModeRadio ? selectedModeRadio.value : 'flashcard';

    // Chuyển màn hình
    document.querySelector('.upload-section').style.display = 'none';
    topicsContainerCard.style.display = 'none';
    studySection.style.display = 'block';

    // Hiển thị nội dung tương ứng chế độ
    modeFlashcardContent.style.display = 'none';
    modeQuizContent.style.display = 'none';
    studyCompleteContent.style.display = 'none';

    if (currentMode === 'flashcard') {
        studyTitle.textContent = '🃏 Thẻ ghi nhớ thông minh';
        modeFlashcardContent.style.display = 'block';
        initFlashcard();
    } else {
        studyTitle.textContent = '📝 Trắc nghiệm từ vựng';
        modeQuizContent.style.display = 'block';
        initQuiz();
    }
});

btnBackToHome.addEventListener('click', () => {
    studySection.style.display = 'none';
    document.querySelector('.upload-section').style.display = 'block';
    updateUI();
});

// Hàm trộn mảng ngẫu nhiên
function shuffleArray(array) {
    for (let i = array.length - 1; i > 0; i--) {
        const j = Math.floor(Math.random() * (i + 1));
        [array[i], array[j]] = [array[j], array[i]];
    }
}

// --- 6. CHẾ ĐỘ 1: FLASHCARD ---
const flashcard = document.getElementById('flashcard');
const fcWord = document.getElementById('fc-word');
const fcMeaning = document.getElementById('fc-meaning');
const fcProgressText = document.getElementById('fc-progress-text');
const fcPrev = document.getElementById('fc-prev');
const fcNext = document.getElementById('fc-next');
const fcAudio = document.getElementById('fc-audio');

function initFlashcard() {
    fcIndex = 0;
    renderFlashcard();
}

function renderFlashcard() {
    if (currentStudyWords.length === 0) return;
    
    // Đảm bảo không bị lật ngược khi đổi thẻ
    flashcard.classList.remove('is-flipped');
    
    const item = currentStudyWords[fcIndex];
    fcWord.textContent = item.word;
    fcMeaning.textContent = item.meaning;
    fcProgressText.textContent = `Thẻ ${fcIndex + 1} / ${currentStudyWords.length}`;
}

flashcard.addEventListener('click', () => {
    flashcard.classList.toggle('is-flipped');
});

fcPrev.addEventListener('click', (e) => {
    e.stopPropagation();
    if (fcIndex > 0) {
        fcIndex--;
        renderFlashcard();
    }
});

fcNext.addEventListener('click', (e) => {
    e.stopPropagation();
    if (fcIndex < currentStudyWords.length - 1) {
        fcIndex++;
        renderFlashcard();
    } else {
        // Hoàn thành flashcard
        modeFlashcardContent.style.display = 'none';
        studyCompleteContent.style.display = 'block';
        document.getElementById('congrats-message').textContent = `Bạn đã xem qua toàn bộ ${currentStudyWords.length} từ vựng!`;
    }
});

fcAudio.addEventListener('click', (e) => {
    e.stopPropagation();
    speakWord(currentStudyWords[fcIndex].word);
});

// --- 7. CHẾ ĐỘ 2: TRẮC NGHIỆM ---
const quizProgress = document.getElementById('quiz-progress');
const quizScoreEl = document.getElementById('quiz-score');
const quizQuestionWord = document.getElementById('quiz-question-word');
const quizOptionsEl = document.getElementById('quiz-options');
const quizNextBtn = document.getElementById('quiz-next-btn');
const quizSpeakBtn = document.getElementById('quiz-speak-btn');

function initQuiz() {
    quizIndex = 0;
    quizScore = 0;
    loadQuizQuestion();
}

function loadQuizQuestion() {
    if (quizIndex >= currentStudyWords.length) {
        // Hoàn thành trắc nghiệm
        modeQuizContent.style.display = 'none';
        studyCompleteContent.style.display = 'block';
        document.getElementById('congrats-message').textContent = `Bạn đạt được ${quizScore} / ${currentStudyWords.length} câu trả lời đúng!`;
        return;
    }

    quizNextBtn.style.display = 'none';
    quizProgress.textContent = `Câu ${quizIndex + 1} / ${currentStudyWords.length}`;
    quizScoreEl.textContent = `⭐ Điểm: ${quizScore}`;

    currentQuizQuestion = currentStudyWords[quizIndex];
    quizQuestionWord.textContent = `Nghĩa của từ "${currentQuizQuestion.word}" là gì?`;

    // Tạo 4 phương án (1 đúng, 3 sai ngẫu nhiên)
    let options = [currentQuizQuestion.meaning];
    
    // Lấy các nghĩa khác làm phương án nhiễu
    let otherWords = currentStudyWords.filter(item => item.meaning !== currentQuizQuestion.meaning);
    shuffleArray(otherWords);
    
    for (let i = 0; i < Math.min(3, otherWords.length); i++) {
        options.push(otherWords[i].meaning);
    }
    shuffleArray(options);

    // Render nút đáp án
    quizOptionsEl.innerHTML = '';
    options.forEach(opt => {
        const btn = document.createElement('button');
        btn.className = 'quiz-option-btn';
        btn.textContent = opt;
        btn.addEventListener('click', () => handleQuizAnswer(btn, opt, currentQuizQuestion.meaning));
        quizOptionsEl.appendChild(btn);
    });

    // Phát âm từ hỏi
    speakWord(currentQuizQuestion.word);
}

quizSpeakBtn.addEventListener('click', () => {
    if (currentQuizQuestion) {
        speakWord(currentQuizQuestion.word);
    }
});

function handleQuizAnswer(selectedBtn, chosenMeaning, correctMeaning) {
    // Vô hiệu hóa tất cả các nút sau khi chọn
    const allOptionBtns = quizOptionsEl.querySelectorAll('.quiz-option-btn');
    allOptionBtns.forEach(btn => btn.disabled = true);

    if (chosenMeaning === correctMeaning) {
        selectedBtn.classList.add('correct');
        quizScore++;
        quizScoreEl.textContent = `⭐ Điểm: ${quizScore}`;
    } else {
        selectedBtn.classList.add('incorrect');
        // Tìm và highlight đáp án đúng
        allOptionBtns.forEach(btn => {
            if (btn.textContent === correctMeaning) {
                btn.classList.add('correct');
            }
        });
    }

    quizNextBtn.style.display = 'block';
}

quizNextBtn.addEventListener('click', () => {
    quizIndex++;
    loadQuizQuestion();
});

// Nút học lại sau khi hoàn thành
document.getElementById('btn-restart-study').addEventListener('click', () => {
    studyCompleteContent.style.display = 'none';
    if (currentMode === 'flashcard') {
        modeFlashcardContent.style.display = 'block';
        initFlashcard();
    } else {
        modeQuizContent.style.display = 'block';
        initQuiz();
    }
});

// --- 8. HỖ TRỢ PHÁT ÂM (SPEECH SYNTHESIS) ---
function speakWord(word) {
    if ('speechSynthesis' in window) {
        window.speechSynthesis.cancel(); // Dừng phát âm cũ nếu đang chạy
        const utterance = new SpeechSynthesisUtterance(word);
        utterance.lang = 'en-US';
        utterance.rate = 0.9; // Đọc chậm một chút cho học sinh dễ nghe
        window.speechSynthesis.speak(utterance);
    }
}