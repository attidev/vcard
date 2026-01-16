let csvData = [];
let qrCodes = [];

// CSV 파일 업로드 처리
document.getElementById('csvFile').addEventListener('change', handleFileSelect);
document.getElementById('uploadBox').addEventListener('dragover', handleDragOver);
document.getElementById('uploadBox').addEventListener('drop', handleDrop);
document.getElementById('uploadBox').addEventListener('dragleave', handleDragLeave);
document.getElementById('generateBtn').addEventListener('click', generateQRCodes);
document.getElementById('downloadAllBtn').addEventListener('click', downloadAll);
document.getElementById('resetBtn').addEventListener('click', reset);

function handleDragOver(e) {
    e.preventDefault();
    e.stopPropagation();
    document.getElementById('uploadBox').classList.add('dragover');
}

function handleDragLeave(e) {
    e.preventDefault();
    e.stopPropagation();
    document.getElementById('uploadBox').classList.remove('dragover');
}

function handleDrop(e) {
    e.preventDefault();
    e.stopPropagation();
    document.getElementById('uploadBox').classList.remove('dragover');
    
    const files = e.dataTransfer.files;
    if (files.length > 0 && files[0].type === 'text/csv' || files[0].name.endsWith('.csv')) {
        processFile(files[0]);
    }
}

function handleFileSelect(e) {
    const file = e.target.files[0];
    if (file) {
        processFile(file);
    }
}

function processFile(file) {
    const reader = new FileReader();
    reader.onload = function(e) {
        const text = e.target.result;
        parseCSV(text);
    };
    reader.readAsText(file, 'UTF-8');
}

function parseCSV(text) {
    const lines = text.split('\n').filter(line => line.trim());
    if (lines.length < 2) {
        showError('CSV 파일에 데이터가 없습니다. 헤더와 최소 1개의 데이터 행이 필요합니다.');
        return;
    }

    const headers = lines[0].split(',').map(h => h.trim());
    csvData = [];

    for (let i = 1; i < lines.length; i++) {
        const values = parseCSVLine(lines[i]);
        if (values.length === 0) continue;
        
        const row = {};
        headers.forEach((header, index) => {
            const value = values[index] ? values[index].trim() : '';
            // 같은 헤더명이 여러 개 있을 경우 배열로 저장 (예: 이메일이 두 개)
            if (row[header] !== undefined) {
                if (Array.isArray(row[header])) {
                    if (value) row[header].push(value);
                } else {
                    const existingValue = row[header];
                    row[header] = existingValue ? [existingValue] : [];
                    if (value) row[header].push(value);
                }
            } else {
                row[header] = value;
            }
        });
        csvData.push(row);
    }

    if (csvData.length === 0) {
        showError('CSV 파일에서 데이터를 읽을 수 없습니다.');
        return;
    }

    document.getElementById('settingsSection').style.display = 'block';
    document.getElementById('uploadBox').style.display = 'none';
    showSuccess(`${csvData.length}명의 데이터를 불러왔습니다.`);
}

function parseCSVLine(line) {
    const result = [];
    let current = '';
    let inQuotes = false;

    for (let i = 0; i < line.length; i++) {
        const char = line[i];
        
        if (char === '"') {
            inQuotes = !inQuotes;
        } else if (char === ',' && !inQuotes) {
            result.push(current);
            current = '';
        } else {
            current += char;
        }
    }
    result.push(current);
    return result;
}

function normalizeText(value) {
    if (!value) return '';
    return value.toString().trim().replace(/\s+/g, ' ');
}

function normalizePhone(value) {
    return normalizeText(value).replace(/[^\d+\-() ]/g, '');
}

// ZXing 방식 참고: QR 코드 용량 계산 (오류 수정 레벨별)
function getMaxCapacity(errorLevel) {
    // 오류 수정 레벨별 최대 바이트 수 (버전 40 기준)
    const capacities = {
        'L': 2953,  // 약 7% 복구 가능
        'M': 2331,  // 약 15% 복구 가능
        'Q': 1663,  // 약 25% 복구 가능
        'H': 1273   // 약 30% 복구 가능
    };
    return capacities[errorLevel] || capacities['L'];
}

// 데이터 길이에 맞는 최적의 오류 수정 레벨 선택
function selectOptimalErrorLevel(byteLength) {
    // ZXing 방식: 데이터가 작으면 높은 오류 수정 레벨 사용 가능
    if (byteLength <= 500) return 'M';   // 중간 크기면 M
    if (byteLength <= 1000) return 'L';  // 크면 L (최대 용량)
    if (byteLength <= 1500) return 'L';  // 매우 크면 L
    return 'L'; // 기본값
}

function generateVCard(row, companyName, companyAddress, companyWebsite, useMinimal = false) {
    const name = row['이름'] || row['name'] || row['Name'] || '';
    const position = row['직책'] || row['position'] || row['Position'] || row['직위'] || '';
    const phone = row['전화번호'] || row['phone'] || row['Phone'] || row['핸드폰'] || '';
    // 이메일이 배열일 수도 있고 단일 값일 수도 있음
    const emailRaw = row['이메일'] || row['email'] || row['Email'] || '';
    const emails = Array.isArray(emailRaw) ? emailRaw : (emailRaw ? [emailRaw] : []);
    const mobile = row['휴대폰'] || row['mobile'] || row['Mobile'] || phone;
    const department = row['부서'] || row['department'] || row['Department'] || '';

    const cleanName = normalizeText(name);
    // 직책은 원본 그대로 유지 (대소문자 보존)
    const cleanPosition = useMinimal ? '' : (position ? position.trim() : '');
    const cleanPhone = normalizePhone(phone);
    const cleanEmails = emails.map(e => normalizeText(e)).filter(e => e);
    const cleanMobile = normalizePhone(mobile);
    const cleanDepartment = normalizeText(department);
    const cleanCompanyName = useMinimal ? '' : normalizeText(companyName);
    const cleanAddress = useMinimal ? '' : normalizeText(companyAddress);
    const cleanWebsite = useMinimal ? '' : normalizeText(companyWebsite);

    // 주소에서 쉼표를 이스케이프 처리 (vCard 표준)
    const escapeAddress = (addr) => {
        if (!addr) return '';
        return addr.replace(/,/g, '\\,').replace(/;/g, '\\;');
    };

    // 최소 버전인 경우 필수 필드만
    if (useMinimal) {
        const vcardParts = [];
        vcardParts.push('BEGIN:VCARD');
        vcardParts.push('VERSION:3.0');
        if (cleanName) {
            vcardParts.push(`N:${cleanName}`); // ZXing 예제: FN 없음
        }
        // 이메일이 여러 개 있을 수 있음
        cleanEmails.forEach(email => {
            if (email) vcardParts.push(`EMAIL:${email}`);
        });
        if (cleanMobile) vcardParts.push(`TEL:${cleanMobile}`);
        vcardParts.push('END:VCARD');
        return vcardParts.join('\n'); // ZXing 방식: \n만 사용
    }

    // vCard 생성 (ZXing 예제 정확히 따름)
    // ZXing 예제: BEGIN:VCARD\nVERSION:3.0\nN:양승대\nORG:남양인터내셔날\nTEL:01081905377\nURL:http://www.namyang-intl.com\nEMAIL:hellojuddy@naver.com\nADR:서울 성동구 뚝섬로1길 63\, 701호\nEND:VCARD
    const vcardParts = [];
    vcardParts.push('BEGIN:VCARD');
    vcardParts.push('VERSION:3.0');
    
    // N 필드: ZXing 예제처럼 단순 (FN 없음!)
    if (cleanName) {
        vcardParts.push(`N:${cleanName}`);
    }
    
    // 회사 정보 (ORG는 N 다음)
    if (cleanCompanyName) {
        vcardParts.push(`ORG:${cleanCompanyName}`);
    }
    
    // 전화번호 (ZXing 예제처럼 단순한 형식, 순서: TEL -> URL -> EMAIL -> ADR)
    if (cleanMobile) {
        vcardParts.push(`TEL:${cleanMobile}`);
    }
    if (cleanPhone && cleanPhone !== cleanMobile) {
        vcardParts.push(`TEL:${cleanPhone}`);
    }
    
    // 웹사이트 (ZXing 예제 순서: TEL 다음)
    if (cleanWebsite) {
        let website = cleanWebsite;
        if (!website.match(/^https?:\/\//i)) {
            website = 'http://' + website;
        }
        vcardParts.push(`URL:${website}`);
    }
    
    // 이메일 (ZXing 예제 순서: URL 다음) - 여러 개의 이메일 지원
    cleanEmails.forEach(email => {
        if (email) vcardParts.push(`EMAIL:${email}`);
    });
    
    // 주소 (ZXing 예제 순서: EMAIL 다음, 쉼표 이스케이프)
    // 예제: ADR:서울 성동구 뚝섬로1길 63\, 701호
    if (cleanAddress) {
        const escapedAddress = escapeAddress(cleanAddress);
        vcardParts.push(`ADR:${escapedAddress}`);
    }
    
    // 직책은 ZXing 예제에 없으므로 마지막에 추가 (선택적)
    if (cleanPosition) {
        vcardParts.push(`TITLE:${cleanPosition}`);
    }
    
    vcardParts.push('END:VCARD');

    // 줄바꿈으로 결합 (ZXing 방식: \n만 사용)
    return vcardParts.join('\n');
}

function buildPayloads(row, companyName, companyAddress, companyWebsite) {
    const payloads = [];
    const fullVCard = generateVCard(row, companyName, companyAddress, companyWebsite, false);
    const minimalVCard = generateVCard(row, companyName, companyAddress, companyWebsite, true);

    // 중요: QR 코드에는 원본 vCard 텍스트를 넣어야 함!
    // URL 인코딩은 ZXing API 서버에 전달할 때만 사용하는 것이지, QR 코드 자체에는 원본 텍스트를 넣어야 함
    
    // 1. 원본 vCard (ZXing 예제 형식으로 생성된 원본 텍스트)
    if (fullVCard) payloads.push(fullVCard);
    if (minimalVCard && minimalVCard !== fullVCard) payloads.push(minimalVCard);

    // 2. 간단한 연락처 방식 (fallback)
    const emailRaw = row['이메일'] || row['email'] || row['Email'] || '';
    const emails = Array.isArray(emailRaw) ? emailRaw : (emailRaw ? [emailRaw] : []);
    const phone = normalizePhone(
        row['휴대폰'] || row['mobile'] || row['Mobile'] || row['전화번호'] || row['phone'] || ''
    );
    const name = normalizeText(row['이름'] || row['name'] || row['Name'] || '');

    if (phone) payloads.push(`tel:${phone}`);
    emails.forEach(email => {
        const cleanEmail = normalizeText(email);
        if (cleanEmail) payloads.push(`mailto:${cleanEmail}`);
    });
    if (name) payloads.push(name);

    return Array.from(new Set(payloads.filter(Boolean)));
}

async function generateQRCodes() {
    // QRCode 라이브러리 로드 확인
    if (typeof QRCode === 'undefined') {
        showError('QRCode 라이브러리를 로드할 수 없습니다. 페이지를 새로고침하거나 인터넷 연결을 확인해주세요.');
        return;
    }

    const companyName = document.getElementById('companyName').value.trim();
    if (!companyName) {
        showError('회사명을 입력해주세요.');
        return;
    }

    const companyAddress = document.getElementById('companyAddress').value.trim();
    const companyWebsite = document.getElementById('companyWebsite').value.trim();

    const qrGrid = document.getElementById('qrGrid');
    const totalCount = csvData.length;
    qrGrid.innerHTML = `<div class="loading">QR 코드 생성 중... (0/${totalCount})</div>`;
    qrCodes = [];

    try {
        for (let i = 0; i < csvData.length; i++) {
            // 진행 상황 업데이트
            if (i % 10 === 0 || i === csvData.length - 1) {
                qrGrid.innerHTML = `<div class="loading">QR 코드 생성 중... (${i + 1}/${totalCount})</div>`;
            }
            const row = csvData[i];
            const payloads = buildPayloads(row, companyName, companyAddress, companyWebsite);
            const name = row['이름'] || row['name'] || row['Name'] || `사원${i + 1}`;
            const position = row['직책'] || row['position'] || row['Position'] || row['직위'] || '';

            // QR 코드 생성 (ZXing 방식: 오류 수정 레벨 자동 선택, 버전 자동)
            const canvas = document.createElement('canvas');
            canvas.width = 300;
            canvas.height = 300;

            let qrCreated = false;
            let lastError = null;
            let payloadUsed = '';

            // QRCode.toCanvas가 있는지 확인 (더 안정적인 방법)
            if (typeof QRCode.toCanvas === 'function') {
                // 간단하고 직접적인 방식: 각 페이로드를 순서대로 시도
                for (const payload of payloads) {
                    try {
                        // 1.2.2 버전 호환: 간단한 옵션만 사용
                        await QRCode.toCanvas(canvas, payload, {
                            width: 300,
                            margin: 1,
                            color: {
                                dark: '#000000',
                                light: '#FFFFFF'
                            },
                            errorCorrectionLevel: 'L'
                        });

                        // 생성 확인: canvas에 실제로 그려졌는지 체크
                        const ctx = canvas.getContext('2d', { willReadFrequently: true });
                        const imageData = ctx.getImageData(0, 0, 10, 10);
                        let hasPixels = false;
                        for (let i = 0; i < imageData.data.length; i += 4) {
                            if (imageData.data[i] !== 255 || imageData.data[i+1] !== 255 || imageData.data[i+2] !== 255) {
                                hasPixels = true;
                                break;
                            }
                        }
                        
                        if (hasPixels) {
                            qrCreated = true;
                            payloadUsed = payload;
                            console.log(`✅ QR 코드 생성 성공: ${name} (${payload.length}자)`);
                            break;
                        } else {
                            throw new Error('Canvas가 비어있습니다');
                        }
                    } catch (error) {
                        lastError = error;
                        console.warn(`QR 코드 생성 실패 (${name}):`, error.message);
                        continue;
                    }
                }

                // 모든 페이로드 실패 시 toDataURL 방식 시도
                if (!qrCreated && payloads.length > 0) {
                    const shortestPayload = payloads[payloads.length - 1];
                    try {
                        const dataUrl = await QRCode.toDataURL(shortestPayload, {
                            width: 300,
                            margin: 1,
                            color: {
                                dark: '#000000',
                                light: '#FFFFFF'
                            },
                            errorCorrectionLevel: 'L'
                        });
                        
                        const img = new Image();
                        await new Promise((resolve, reject) => {
                            const timeout = setTimeout(() => reject(new Error('이미지 로딩 타임아웃')), 3000);
                            img.onload = () => {
                                clearTimeout(timeout);
                                const ctx = canvas.getContext('2d', { willReadFrequently: true });
                                ctx.clearRect(0, 0, 300, 300);
                                ctx.drawImage(img, 0, 0, 300, 300);
                                qrCreated = true;
                                payloadUsed = shortestPayload;
                                resolve();
                            };
                            img.onerror = () => {
                                clearTimeout(timeout);
                                reject(new Error('이미지 로딩 실패'));
                            };
                            img.src = dataUrl;
                        });
                        console.log(`✅ QR 코드 생성 성공 (toDataURL): ${name}`);
                    } catch (error) {
                        lastError = lastError || error;
                        console.warn(`toDataURL 방식도 실패:`, error.message);
                    }
                }

                // 최후의 수단: 이름만
                if (!qrCreated) {
                    const fallbackText = name || 'QR';
                    try {
                        await QRCode.toCanvas(canvas, fallbackText, {
                            width: 300,
                            margin: 1,
                            color: { dark: '#000000', light: '#FFFFFF' },
                            errorCorrectionLevel: 'L'
                        });
                        qrCreated = true;
                        payloadUsed = fallbackText;
                        console.log(`✅ QR 코드 생성 성공 (fallback): ${name}`);
                    } catch (error) {
                        lastError = lastError || error;
                        console.error(`Fallback도 실패:`, error.message);
                    }
                }
            } else {
                // 구버전 API 사용 (DOM 요소 방식) - 간단하게
                const tempDiv = document.createElement('div');
                tempDiv.style.width = '300px';
                tempDiv.style.height = '300px';
                tempDiv.style.position = 'absolute';
                tempDiv.style.left = '-9999px';
                document.body.appendChild(tempDiv);

                for (const payload of payloads) {
                    try {
                        tempDiv.innerHTML = '';
                        new QRCode(tempDiv, {
                            text: payload,
                            width: 300,
                            height: 300,
                            colorDark: '#000000',
                            colorLight: '#FFFFFF',
                            correctLevel: QRCode.CorrectLevel.L
                        });

                        // 충분한 시간 대기
                        await new Promise(resolve => setTimeout(resolve, 800));

                        const img = tempDiv.querySelector('img');
                        const qrCanvas = tempDiv.querySelector('canvas');

                        if (qrCanvas) {
                            const ctx = canvas.getContext('2d', { willReadFrequently: true });
                            ctx.clearRect(0, 0, 300, 300);
                            ctx.drawImage(qrCanvas, 0, 0, 300, 300);
                            qrCreated = true;
                            payloadUsed = payload;
                            console.log(`✅ QR 코드 생성 성공 (legacy canvas): ${name}`);
                            break;
                        } else if (img && img.src) {
                            // 이미지 로딩 대기
                            await new Promise((resolve, reject) => {
                                const timeout = setTimeout(() => reject(new Error('타임아웃')), 5000);
                                const checkImage = () => {
                                    if (img.complete && img.naturalWidth > 0) {
                                        clearTimeout(timeout);
                                        const ctx = canvas.getContext('2d', { willReadFrequently: true });
                                        ctx.clearRect(0, 0, 300, 300);
                                        ctx.drawImage(img, 0, 0, 300, 300);
                                        qrCreated = true;
                                        payloadUsed = payload;
                                        resolve();
                                    }
                                };
                                img.onload = checkImage;
                                img.onerror = () => {
                                    clearTimeout(timeout);
                                    reject(new Error('이미지 로딩 실패'));
                                };
                                checkImage(); // 즉시 확인
                            });
                            if (qrCreated) {
                                console.log(`✅ QR 코드 생성 성공 (legacy img): ${name}`);
                                break;
                            }
                        }
                    } catch (error) {
                        lastError = error;
                        console.warn(`QR 코드 생성 실패 (legacy, ${name}):`, error.message);
                        continue;
                    }
                }

                // 최후의 수단
                if (!qrCreated) {
                    try {
                        tempDiv.innerHTML = '';
                        const fallbackText = name || 'QR';
                        new QRCode(tempDiv, {
                            text: fallbackText,
                            width: 300,
                            height: 300,
                            colorDark: '#000000',
                            colorLight: '#FFFFFF',
                            correctLevel: QRCode.CorrectLevel.L
                        });
                        
                        await new Promise(resolve => setTimeout(resolve, 800));
                        
                        const img = tempDiv.querySelector('img');
                        const qrCanvas = tempDiv.querySelector('canvas');
                        
                        if (qrCanvas) {
                            const ctx = canvas.getContext('2d', { willReadFrequently: true });
                            ctx.clearRect(0, 0, 300, 300);
                            ctx.drawImage(qrCanvas, 0, 0, 300, 300);
                            qrCreated = true;
                            payloadUsed = fallbackText;
                            console.log(`✅ QR 코드 생성 성공 (fallback): ${name}`);
                        } else if (img && img.src && img.complete) {
                            const ctx = canvas.getContext('2d', { willReadFrequently: true });
                            ctx.clearRect(0, 0, 300, 300);
                            ctx.drawImage(img, 0, 0, 300, 300);
                            qrCreated = true;
                            payloadUsed = fallbackText;
                            console.log(`✅ QR 코드 생성 성공 (fallback img): ${name}`);
                        }
                    } catch (error) {
                        lastError = lastError || error;
                        console.error(`Fallback도 실패:`, error.message);
                    }
                }

                document.body.removeChild(tempDiv);
            }
            
            if (!qrCreated) {
                const errorDetails = lastError 
                    ? `오류: ${lastError.message}` 
                    : '모든 페이로드와 오류 수정 레벨에서 실패했습니다.';
                console.error(`QR 코드 생성 실패 (${name}):`, errorDetails);
                console.error('시도한 페이로드:', payloads);
                
                // 오류를 던지지 않고 빈 canvas로 진행 (사용자 경험 개선)
                // throw new Error(`QR 코드 생성 실패: ${errorDetails}`);
            }
            
            // canvas 검증: 실제로 QR 코드가 그려졌는지 확인
            if (qrCreated) {
                const ctx = canvas.getContext('2d', { willReadFrequently: true });
                const imageData = ctx.getImageData(0, 0, canvas.width, canvas.height);
                let blackPixels = 0;
                let whitePixels = 0;
                
                for (let i = 0; i < imageData.data.length; i += 4) {
                    const r = imageData.data[i];
                    const g = imageData.data[i + 1];
                    const b = imageData.data[i + 2];
                    const a = imageData.data[i + 3];
                    
                    if (a > 0) {
                        if (r < 128 || g < 128 || b < 128) {
                            blackPixels++;
                        } else {
                            whitePixels++;
                        }
                    }
                }
                
                // 검은색과 흰색 픽셀이 모두 있어야 QR 코드
                if (blackPixels === 0 || whitePixels === 0) {
                    console.warn(`QR 코드 ${i + 1} (${name})의 canvas가 비어있습니다. (검은색: ${blackPixels}, 흰색: ${whitePixels})`);
                    qrCreated = false;
                } else {
                    console.log(`QR 코드 검증 성공: ${name} (검은색: ${blackPixels}, 흰색: ${whitePixels})`);
                }
            }
            
            // QR 코드가 생성되었거나 실패했어도 항상 추가
            qrCodes.push({
                canvas: canvas,
                name: name,
                position: position,
                vcard: payloadUsed || payloads[0] || '',
                success: qrCreated
            });
        }

        displayQRCodes();
        document.getElementById('previewSection').style.display = 'block';
    } catch (error) {
        showError('QR 코드 생성 중 오류가 발생했습니다: ' + error.message);
        console.error('QR 코드 생성 오류:', error);
    }
}

function displayQRCodes() {
    const qrGrid = document.getElementById('qrGrid');
    qrGrid.innerHTML = '';

    if (qrCodes.length === 0) {
        qrGrid.innerHTML = '<div class="error">생성된 QR 코드가 없습니다.</div>';
        return;
    }

    qrCodes.forEach((qr, index) => {
        const item = document.createElement('div');
        item.className = 'qr-item';
        
        // canvas 복사: cloneNode는 내용을 복사하지 않으므로 새 canvas에 그리기
        const displayCanvas = document.createElement('canvas');
        displayCanvas.width = qr.canvas.width;
        displayCanvas.height = qr.canvas.height;
        displayCanvas.style.display = 'block';
        displayCanvas.style.maxWidth = '100%';
        displayCanvas.style.height = 'auto';
        displayCanvas.style.border = '1px solid #ddd';
        displayCanvas.style.borderRadius = '8px';
        displayCanvas.style.background = '#fff';
        
        // 원본 canvas의 내용을 새 canvas에 그리기
        const ctx = displayCanvas.getContext('2d', { willReadFrequently: true });
        ctx.drawImage(qr.canvas, 0, 0);
        
        // 실패한 QR 코드 표시
        if (qr.success === false) {
            const errorMsg = document.createElement('div');
            errorMsg.className = 'error';
            errorMsg.style.cssText = 'padding: 10px; background: #fee; color: #c33; border-radius: 4px; margin-bottom: 10px; font-size: 12px;';
            errorMsg.textContent = '⚠️ QR 코드 생성 실패';
            item.appendChild(errorMsg);
        }
        
        const nameDiv = document.createElement('div');
        nameDiv.className = 'name';
        nameDiv.textContent = qr.name;
        
        const positionDiv = document.createElement('div');
        positionDiv.className = 'position';
        positionDiv.textContent = qr.position || '';

        const downloadBtn = document.createElement('button');
        downloadBtn.className = 'download-btn';
        downloadBtn.textContent = '다운로드';
        downloadBtn.onclick = () => downloadQR(index);
        if (qr.success === false) {
            downloadBtn.disabled = true;
            downloadBtn.style.opacity = '0.5';
            downloadBtn.style.cursor = 'not-allowed';
        }

        item.appendChild(displayCanvas);
        item.appendChild(nameDiv);
        if (qr.position) item.appendChild(positionDiv);
        item.appendChild(downloadBtn);
        
        qrGrid.appendChild(item);
    });
}

function downloadQR(index) {
    const qr = qrCodes[index];
    const name = qr.name.replace(/[^a-zA-Z0-9가-힣]/g, '_');
    
    qr.canvas.toBlob((blob) => {
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `${name}_QR코드.png`;
        a.click();
        URL.revokeObjectURL(url);
    });
}

async function downloadAll() {
    if (qrCodes.length === 0) return;

    const zip = new JSZip();
    const loading = document.createElement('div');
    loading.className = 'loading';
    loading.textContent = 'ZIP 파일 생성 중...';
    document.getElementById('qrGrid').prepend(loading);

    try {
        for (let i = 0; i < qrCodes.length; i++) {
            const qr = qrCodes[i];
            const name = qr.name.replace(/[^a-zA-Z0-9가-힣]/g, '_');
            
            const blob = await new Promise(resolve => {
                qr.canvas.toBlob(resolve);
            });
            
            zip.file(`${name}_QR코드.png`, blob);
        }

        const zipBlob = await zip.generateAsync({ type: 'blob' });
        const url = URL.createObjectURL(zipBlob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `QR코드_전체_${new Date().getTime()}.zip`;
        a.click();
        URL.revokeObjectURL(url);
        
        loading.remove();
    } catch (error) {
        loading.remove();
        showError('ZIP 파일 생성 중 오류가 발생했습니다: ' + error.message);
    }
}

function reset() {
    csvData = [];
    qrCodes = [];
    document.getElementById('csvFile').value = '';
    document.getElementById('companyName').value = '';
    document.getElementById('companyAddress').value = '';
    document.getElementById('companyWebsite').value = '';
    document.getElementById('settingsSection').style.display = 'none';
    document.getElementById('previewSection').style.display = 'none';
    document.getElementById('uploadBox').style.display = 'block';
    document.getElementById('qrGrid').innerHTML = '';
}

function showError(message) {
    const errorDiv = document.createElement('div');
    errorDiv.className = 'error';
    errorDiv.textContent = message;
    document.querySelector('.container').insertBefore(errorDiv, document.querySelector('.container').firstChild);
    setTimeout(() => errorDiv.remove(), 5000);
}

function showSuccess(message) {
    const successDiv = document.createElement('div');
    successDiv.style.cssText = 'background: #efe; color: #3c3; padding: 15px; border-radius: 8px; margin: 20px 0; border: 2px solid #cfc;';
    successDiv.textContent = message;
    document.querySelector('.container').insertBefore(successDiv, document.querySelector('.container').firstChild);
    setTimeout(() => successDiv.remove(), 3000);
}
