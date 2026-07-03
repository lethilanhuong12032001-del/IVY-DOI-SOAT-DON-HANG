document.addEventListener('DOMContentLoaded', () => {
    // Application State (Runs 100% Client-Side)
    let state = {
        files: {
            internal: false,
            tiktok_all: false,
            tiktok_returns: false
        },
        data: {
            internal: null,
            tiktok_all: null,
            tiktok_returns: null
        },
        reconciledData: null,
        activeTab: 'upload-tab'
    };

    // UI Elements
    const tabs = document.querySelectorAll('.tab-btn');
    const tabContents = document.querySelectorAll('.tab-content');
    const dropzones = {
        internal: document.getElementById('dropzone-internal'),
        tiktok_all: document.getElementById('dropzone-tiktok-all'),
        tiktok_returns: document.getElementById('dropzone-tiktok-returns')
    };
    const fileInputs = {
        internal: document.getElementById('file-internal'),
        tiktok_all: document.getElementById('file-tiktok-all'),
        tiktok_returns: document.getElementById('file-tiktok-returns')
    };
    const infoBars = {
        internal: document.getElementById('info-internal'),
        tiktok_all: document.getElementById('info-tiktok-all'),
        tiktok_returns: document.getElementById('info-tiktok-returns')
    };
    const runBtn = document.getElementById('run-reconcile-btn');
    const clearBtn = document.getElementById('clear-all-btn');
    
    // Overview elements
    const statTotalTiktok = document.getElementById('stat-total-tiktok');
    const statTotalInternal = document.getElementById('stat-total-internal');
    const statTotalReturns = document.getElementById('stat-total-returns');
    const statCase1Count = document.getElementById('stat-case1-count');
    const statCase2Count = document.getElementById('stat-case2-count');
    const statLossRate = document.getElementById('stat-loss-rate');
    const gaugeFill = document.getElementById('gauge-fill-element');
    const gaugePercentageText = document.getElementById('gauge-percentage-text');

    // Case 1 Table elements
    const case1Search = document.getElementById('case1-search');
    const case1ReasonFilter = document.getElementById('case1-reason-filter');
    const case1Tbody = document.getElementById('case1-tbody');
    const case1PaginationInfo = document.getElementById('case1-pagination-info');
    const exportCase1Btn = document.getElementById('export-case1-btn');

    // Case 2 Table elements
    const case2Search = document.getElementById('case2-search');
    const case2CarrierFilter = document.getElementById('case2-carrier-filter');
    const case2Tbody = document.getElementById('case2-tbody');
    const case2PaginationInfo = document.getElementById('case2-pagination-info');
    const exportCase2Btn = document.getElementById('export-case2-btn');

    // Loader overlay & Modal
    const loaderOverlay = document.getElementById('loader-overlay');
    const errorModal = document.getElementById('error-modal');
    const errorModalMessage = document.getElementById('error-modal-message');
    const closeModalBtns = [
        document.getElementById('close-modal-btn'),
        document.getElementById('close-modal-footer-btn')
    ];

    // 1. Tab Navigation Handling
    tabs.forEach(tab => {
        tab.addEventListener('click', () => {
            if (tab.disabled) return;
            
            // Toggle active tab button
            tabs.forEach(t => t.classList.remove('active'));
            tab.classList.add('active');
            
            // Toggle active tab content
            const targetTab = tab.getAttribute('data-tab');
            tabContents.forEach(content => {
                if (content.id === targetTab) {
                    content.classList.add('active');
                } else {
                    content.classList.remove('active');
                }
            });
            state.activeTab = targetTab;
        });
    });

    // 2. Drag and Drop File Handlers
    Object.keys(dropzones).forEach(key => {
        const dropzone = dropzones[key];
        const fileInput = fileInputs[key];

        // Highlight drop zone on drag over
        ['dragenter', 'dragover'].forEach(eventName => {
            dropzone.addEventListener(eventName, (e) => {
                e.preventDefault();
                e.stopPropagation();
                dropzone.classList.add('dragover');
            }, false);
        });

        ['dragleave', 'drop'].forEach(eventName => {
            dropzone.addEventListener(eventName, (e) => {
                e.preventDefault();
                e.stopPropagation();
                dropzone.classList.remove('dragover');
            }, false);
        });

        // Handle drop event
        dropzone.addEventListener('drop', (e) => {
            const dt = e.dataTransfer;
            const file = dt.files[0];
            if (file) {
                handleFileUpload(file, key);
            }
        });

        // Handle file input selection
        fileInput.addEventListener('change', (e) => {
            const file = e.target.files[0];
            if (file) {
                handleFileUpload(file, key);
                fileInput.value = ''; // Reset value so change event can re-fire
            }
        });
    });

    // 3. Client-Side File Reading & Parsing
    function handleFileUpload(file, key) {
        // Validate extension
        const isCsv = file.name.toLowerCase().endsWith('.csv');
        const isXlsx = file.name.toLowerCase().endsWith('.xlsx');
        
        if (key === 'internal' && !isCsv) {
            showError("Vui lòng tải lên file định dạng CSV (.csv) cho dữ liệu nội bộ kho.");
            return;
        }
        if (key !== 'internal' && !isXlsx) {
            showError("Vui lòng tải lên file định dạng Excel (.xlsx) cho báo cáo TikTok.");
            return;
        }

        // Show file info bar with loading status
        const infoBar = infoBars[key];
        infoBar.classList.remove('hidden');
        infoBar.querySelector('.file-name').textContent = file.name;
        infoBar.querySelector('.file-size').textContent = formatBytes(file.size);
        
        const badge = infoBar.querySelector('.status-badge');
        badge.className = 'status-badge status-loading';
        badge.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> Đang đọc file...';

        if (key === 'internal') {
            // Read CSV File
            readCSVFile(file, (text) => {
                try {
                    const parsedData = parseCSV(text);
                    if (parsedData.length === 0) {
                        throw new Error("File CSV trống hoặc không đúng cấu trúc.");
                    }
                    state.data.internal = parsedData;
                    state.files.internal = true;
                    
                    badge.className = 'status-badge status-success';
                    badge.innerHTML = `<i class="fa-solid fa-check"></i> Đã đọc ${formatNumber(parsedData.length)} dòng`;
                    checkReconcileAvailability();
                } catch (err) {
                    badge.className = 'status-badge status-error';
                    badge.innerHTML = '<i class="fa-solid fa-xmark"></i> Lỗi đọc file';
                    state.files.internal = false;
                    state.data.internal = null;
                    checkReconcileAvailability();
                    showError("Lỗi phân tích CSV: " + err.message);
                }
            });
        } else {
            // Read Excel File
            readExcelFile(file, (rows) => {
                try {
                    let parsedData = [];
                    if (key === 'tiktok_all') {
                        parsedData = parseTikTokAll(rows);
                    } else {
                        parsedData = parseTikTokReturns(rows);
                    }

                    if (parsedData.length === 0) {
                        throw new Error("File Excel trống hoặc không đúng cấu trúc.");
                    }

                    state.data[key] = parsedData;
                    state.files[key] = true;
                    
                    badge.className = 'status-badge status-success';
                    badge.innerHTML = `<i class="fa-solid fa-check"></i> Đã đọc ${formatNumber(parsedData.length)} dòng`;
                    checkReconcileAvailability();
                } catch (err) {
                    badge.className = 'status-badge status-error';
                    badge.innerHTML = '<i class="fa-solid fa-xmark"></i> Lỗi đọc file';
                    state.files[key] = false;
                    state.data[key] = null;
                    checkReconcileAvailability();
                    showError("Lỗi phân tích Excel: " + err.message);
                }
            });
        }
    }

    // Read CSV with auto encoding detection (UTF-8 / UTF-16)
    function readCSVFile(file, callback) {
        const reader = new FileReader();
        reader.onload = function(e) {
            const buffer = e.target.result;
            const view = new DataView(buffer);
            let encoding = 'utf-8';
            if (view.byteLength >= 2) {
                const bom = view.getUint16(0, true);
                if (bom === 0xFEFF || bom === 0xFFFE) {
                    encoding = 'utf-16le';
                }
            }
            const decoder = new TextDecoder(encoding);
            const text = decoder.decode(buffer);
            callback(text);
        };
        reader.readAsArrayBuffer(file);
    }

    // Parse CSV with auto separator detection
    function parseCSV(text) {
        const lines = text.split(/\r?\n/);
        if (lines.length === 0) return [];
        
        const firstLine = lines[0];
        let sep = '\t';
        if (firstLine.includes('\t')) {
            sep = '\t';
        } else if (firstLine.includes(';')) {
            sep = ';';
        } else if (firstLine.includes(',')) {
            sep = ',';
        }
        
        const headers = firstLine.split(sep).map(h => h.replace(/^"|"$/g, '').trim());
        const data = [];
        
        for (let i = 1; i < lines.length; i++) {
            const line = lines[i].trim();
            if (!line) continue;
            
            const cells = line.split(sep).map(c => c.replace(/^"|"$/g, '').trim());
            const row = {};
            headers.forEach((h, idx) => {
                row[h] = cells[idx] || '';
            });
            data.push(row);
        }
        return data;
    }

    // Read Excel File using SheetJS
    function readExcelFile(file, callback) {
        const reader = new FileReader();
        reader.onload = function(e) {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const sheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[sheetName];
            const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
            callback(rows);
        };
        reader.readAsArrayBuffer(file);
    }

    // Parse TikTok All Orders (skip description row 1)
    function parseTikTokAll(rows) {
        if (rows.length < 3) return [];
        const headers = rows[0].map(h => String(h || '').trim());
        const data = [];
        for (let i = 2; i < rows.length; i++) {
            const row = rows[i];
            const item = {};
            headers.forEach((h, idx) => {
                item[h] = row[idx] !== undefined && row[idx] !== null ? row[idx] : '';
            });
            data.push(item);
        }
        return data;
    }

    // Parse TikTok Returns File
    function parseTikTokReturns(rows) {
        if (rows.length < 2) return [];
        const headers = rows[0].map(h => String(h || '').trim());
        const data = [];
        for (let i = 1; i < rows.length; i++) {
            const row = rows[i];
            const item = {};
            headers.forEach((h, idx) => {
                item[h] = row[idx] !== undefined && row[idx] !== null ? row[idx] : '';
            });
            data.push(item);
        }
        return data;
    }

    function checkReconcileAvailability() {
        const allUploaded = state.files.internal && state.files.tiktok_all && state.files.tiktok_returns;
        runBtn.disabled = !allUploaded;
    }

    // 4. Clear Files Handler
    clearBtn.addEventListener('click', () => {
        // Reset state
        state.files = { internal: false, tiktok_all: false, tiktok_returns: false };
        state.data = { internal: null, tiktok_all: null, tiktok_returns: null };
        state.reconciledData = null;
        
        // Reset UI
        Object.keys(infoBars).forEach(key => {
            infoBars[key].classList.add('hidden');
        });
        
        runBtn.disabled = true;
        
        // Disable navigation
        document.getElementById('nav-overview-btn').disabled = true;
        document.getElementById('nav-case1-btn').disabled = true;
        document.getElementById('nav-case2-btn').disabled = true;

        // Reset file inputs values
        Object.keys(fileInputs).forEach(k => {
            fileInputs[k].value = '';
        });
        
        alert("Đã xóa sạch bộ nhớ tạm!");
    });

    // 5. Client-Side Reconciliation Engine (No Server Call)
    runBtn.addEventListener('click', () => {
        loaderOverlay.classList.remove('hidden');
        
        // Run in timeout to let browser show loader
        setTimeout(() => {
            try {
                // 1. Process Internal Data
                const df_internal = state.data.internal;
                const internal_data = {};
                
                df_internal.forEach(row => {
                    const orderId = String(row['Mã đơn hàng'] || '').trim();
                    const status = String(row['Trạng thái IVY'] || '').trim();
                    const tracking = String(row['Mã vận đơn'] || '').trim();
                    const ivyCode = String(row['Mã IVY'] || '').trim();
                    
                    if (!orderId) return;
                    
                    if (!internal_data[orderId]) {
                        internal_data[orderId] = {
                            statuses: new Set([status]),
                            tracking: tracking ? new Set([tracking]) : new Set(),
                            ivyCodes: ivyCode ? new Set([ivyCode]) : new Set()
                        };
                    } else {
                        internal_data[orderId].statuses.add(status);
                        if (tracking) internal_data[orderId].tracking.add(tracking);
                        if (ivyCode) internal_data[orderId].ivyCodes.add(ivyCode);
                    }
                });
                
                // Convert sets to string
                Object.keys(internal_data).forEach(oid => {
                    internal_data[oid].statuses_str = Array.from(internal_data[oid].statuses).join(', ');
                    internal_data[oid].tracking_str = Array.from(internal_data[oid].tracking).join(', ');
                    internal_data[oid].ivy_codes_str = Array.from(internal_data[oid].ivyCodes).join(', ');
                });

                // 2. Perform Case 1 Reconciliation: Completed Returns on TikTok but missing/unreturned in Warehouse
                const df_tiktok_ret = state.data.tiktok_returns;
                const case1_list = [];
                const seenCase1 = new Set();
                
                df_tiktok_ret.forEach(row => {
                    const returnStatus = String(row['Return Status'] || '').trim();
                    if (returnStatus !== 'Completed') return;
                    
                    const orderId = String(row['Order ID'] || '').trim();
                    if (!orderId) return;
                    if (seenCase1.has(orderId)) return;
                    seenCase1.add(orderId);
                    
                    const sellerSku = String(row['Seller SKU'] || '').trim();
                    const buyer = String(row['Buyer Username'] || '').trim();
                    const reason = String(row['Return Reason'] || '').trim();
                    const refundTime = String(row['Refund Time'] || '').trim();
                    const trackingRet = String(row['Return Logistics Tracking ID'] || '').trim();
                    
                    const internalInfo = internal_data[orderId];
                    if (!internalInfo) {
                        case1_list.push({
                            order_id: orderId,
                            product_code: sellerSku,
                            buyer: buyer,
                            reason: reason,
                            refund_time: refundTime,
                            tracking_id: trackingRet,
                            internal_status: "KHÔNG TÌM THẤY TRÊN HỆ THỐNG",
                            internal_tracking: "",
                            ivy_code: ""
                        });
                    } else {
                        const statuses = internalInfo.statuses;
                        const hasReturnStatus = Array.from(statuses).some(s => s === "Đã Trả hàng/Hủy COD" || s === "Đã Trả hàng/Đổi hàng");
                        if (!hasReturnStatus) {
                            case1_list.push({
                                order_id: orderId,
                                product_code: sellerSku,
                                buyer: buyer,
                                reason: reason,
                                refund_time: refundTime,
                                tracking_id: trackingRet,
                                internal_status: internalInfo.statuses_str,
                                internal_tracking: internalInfo.tracking_str,
                                ivy_code: internalInfo.ivy_codes_str
                            });
                        }
                    }
                });

                // 3. Perform Case 2 Reconciliation: Cancelled or transit return on TikTok, but still "Chờ giao vận" internally
                const df_tiktok_all = state.data.tiktok_all;
                const case2_list = [];
                const seenCase2 = new Set();
                
                df_tiktok_all.forEach(row => {
                    const orderId = String(row['Order ID'] || '').trim();
                    const orderStatus = String(row['Order Status'] || '').trim();
                    const cancelType = String(row['Cancelation/Return Type'] || '').trim();
                    
                    if (orderStatus !== 'Đã hủy' && cancelType !== 'Return/Refund') return;
                    if (!orderId) return;
                    if (seenCase2.has(orderId)) return;
                    seenCase2.add(orderId);
                    
                    const cancelReason = String(row['Cancel Reason'] || '').trim();
                    const recipient = String(row['Recipient'] || '').trim();
                    const phone = String(row['Phone #'] || '').trim();
                    const trackingId = String(row['Tracking ID'] || '').trim();
                    const carrier = String(row['Shipping Provider Name'] || '').trim();
                    
                    const internalInfo = internal_data[orderId];
                    if (!internalInfo) return; // Not warehouse mismatch if not shipped/packed
                    
                    const statuses = internalInfo.statuses;
                    const hasShipped = statuses.has("Chờ giao vận");
                    const hasReturnedOrCancelled = Array.from(statuses).some(s => s === "Đã Trả hàng/Hủy COD" || s === "Đã Trả hàng/Đổi hàng" || s === "Đã hủy đơn hàng");
                    
                    if (hasShipped && !hasReturnedOrCancelled) {
                        case2_list.push({
                            order_id: orderId,
                            platform_status: `${orderStatus} (${cancelType})`,
                            cancel_reason: cancelReason,
                            recipient: recipient,
                            phone: phone,
                            tracking_id: trackingId,
                            carrier: carrier,
                            internal_status: internalInfo.statuses_str,
                            internal_tracking: internalInfo.tracking_str,
                            ivy_code: internalInfo.ivy_codes_str
                        });
                    }
                });

                // Compile summary stats
                const totalTiktok = new Set(df_tiktok_all.map(r => r['Order ID']).filter(Boolean)).size;
                const totalReturns = new Set(df_tiktok_ret.map(r => r['Order ID']).filter(Boolean)).size;
                const totalInternal = Object.keys(internal_data).length;
                
                const summary = {
                    total_tiktok_orders: totalTiktok,
                    total_tiktok_returns: totalReturns,
                    total_internal_orders: totalInternal,
                    case1_count: case1_list.length,
                    case2_count: case2_list.length,
                    total_discrepancies: case1_list.length + case2_list.length
                };
                
                state.reconciledData = {
                    success: true,
                    summary: summary,
                    case1: case1_list,
                    case2: case2_list
                };

                // Unlock tabs
                document.getElementById('nav-overview-btn').disabled = false;
                document.getElementById('nav-case1-btn').disabled = false;
                document.getElementById('nav-case2-btn').disabled = false;
                
                // Update Badge Counts
                document.getElementById('case1-count-badge').textContent = summary.case1_count;
                document.getElementById('case2-count-badge').textContent = summary.case2_count;
                
                // Render Statistics
                renderOverview(summary);
                
                // Populate Filters & Tables
                initializeCase1Filters(case1_list);
                renderCase1Table(case1_list);
                
                initializeCase2Filters(case2_list);
                renderCase2Table(case2_list);
                
                loaderOverlay.classList.add('hidden');
                
                // Switch to Overview tab
                document.getElementById('nav-overview-btn').click();
            } catch (err) {
                loaderOverlay.classList.add('hidden');
                showError("Đã xảy ra lỗi khi đối soát dữ liệu: <br>" + err.message);
            }
        }, 100);
    });

    // 6. Render Statistics Dashboard
    function renderOverview(summary) {
        statTotalTiktok.textContent = formatNumber(summary.total_tiktok_orders);
        statTotalInternal.textContent = formatNumber(summary.total_internal_orders);
        statTotalReturns.textContent = formatNumber(summary.total_tiktok_returns);
        statCase1Count.textContent = formatNumber(summary.case1_count);
        statCase2Count.textContent = formatNumber(summary.case2_count);
        
        const totalDiscrepancies = summary.case1_count + summary.case2_count;
        const totalAffected = summary.total_tiktok_returns + summary.case2_count;
        
        let matchPercentage = 100;
        if (totalAffected > 0) {
            matchPercentage = Math.round(((totalAffected - totalDiscrepancies) / totalAffected) * 100);
            matchPercentage = Math.max(0, Math.min(100, matchPercentage));
        }
        
        statLossRate.textContent = `${100 - matchPercentage}%`;
        if (100 - matchPercentage > 5) {
            statLossRate.className = 'stat-value text-danger';
        } else {
            statLossRate.className = 'stat-value text-warning';
        }
        
        gaugePercentageText.textContent = `${matchPercentage}%`;
        gaugeFill.style.setProperty('--gauge-fill-percentage', `${matchPercentage}%`);
        
        if (matchPercentage > 90) {
            gaugeFill.style.background = `conic-gradient(var(--success) ${matchPercentage}%, #1e293b ${matchPercentage}%)`;
        } else if (matchPercentage > 75) {
            gaugeFill.style.background = `conic-gradient(var(--warning) ${matchPercentage}%, #1e293b ${matchPercentage}%)`;
        } else {
            gaugeFill.style.background = `conic-gradient(var(--danger) ${matchPercentage}%, #1e293b ${matchPercentage}%)`;
        }
    }

    // 7. Case 1: Filter & Search setup
    function initializeCase1Filters(case1Data) {
        const reasons = new Set();
        case1Data.forEach(item => {
            if (item.reason) reasons.add(item.reason);
        });
        
        case1ReasonFilter.innerHTML = '<option value="">Tất cả lý do trả hàng</option>';
        Array.from(reasons).sort().forEach(reason => {
            const opt = document.createElement('option');
            opt.value = reason;
            opt.textContent = reason;
            case1ReasonFilter.appendChild(opt);
        });

        case1Search.oninput = () => filterCase1();
        case1ReasonFilter.onchange = () => filterCase1();
    }

    function filterCase1() {
        if (!state.reconciledData) return;
        
        const q = case1Search.value.toLowerCase().trim();
        const selectedReason = case1ReasonFilter.value;
        
        const filtered = state.reconciledData.case1.filter(item => {
            const matchesQuery = !q || 
                item.order_id.toLowerCase().includes(q) ||
                (item.product_code && item.product_code.toLowerCase().includes(q)) ||
                (item.buyer && item.buyer.toLowerCase().includes(q)) ||
                (item.ivy_code && item.ivy_code.toLowerCase().includes(q));
                
            const matchesReason = !selectedReason || item.reason === selectedReason;
            
            return matchesQuery && matchesReason;
        });
        
        renderCase1Table(filtered);
    }

    function renderCase1Table(data) {
        case1Tbody.innerHTML = '';
        
        if (data.length === 0) {
            case1Tbody.innerHTML = '<tr><td colspan="10" class="text-center">Không tìm thấy đơn hàng lệch khớp nào.</td></tr>';
            case1PaginationInfo.textContent = 'Đang hiển thị 0/0 đơn hàng bị lệch.';
            return;
        }

        data.forEach((item, idx) => {
            const tr = document.createElement('tr');
            const internalBadge = item.internal_status === "KHÔNG TÌM THẤY TRÊN HỆ THỐNG" 
                ? `<span class="badge-status badge-status-danger">${item.internal_status}</span>`
                : `<span class="badge-status badge-status-warning">${item.internal_status}</span>`;
                
            tr.innerHTML = `
                <td>${idx + 1}</td>
                <td class="font-semibold text-primary">${item.order_id}</td>
                <td>${item.product_code || '-'}</td>
                <td class="text-indigo">${item.ivy_code || '-'}</td>
                <td>${item.buyer}</td>
                <td style="max-width: 250px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;" title="${item.reason}">${item.reason}</td>
                <td><span class="badge-status badge-status-success">Trả hàng hoàn tất</span></td>
                <td>${internalBadge}</td>
                <td>${item.refund_time || '-'}</td>
                <td>${item.tracking_id || '-'}</td>
            `;
            case1Tbody.appendChild(tr);
        });

        case1PaginationInfo.textContent = `Đang hiển thị ${data.length}/${state.reconciledData.case1.length} đơn hàng bị lệch.`;
    }

    // 8. Case 2: Filter & Search setup
    function initializeCase2Filters(case2Data) {
        const carriers = new Set();
        case2Data.forEach(item => {
            if (item.carrier) carriers.add(item.carrier);
        });
        
        case2CarrierFilter.innerHTML = '<option value="">Tất cả đơn vị vận chuyển</option>';
        Array.from(carriers).sort().forEach(carrier => {
            const opt = document.createElement('option');
            opt.value = carrier;
            opt.textContent = carrier;
            case2CarrierFilter.appendChild(opt);
        });

        case2Search.oninput = () => filterCase2();
        case2CarrierFilter.onchange = () => filterCase2();
    }

    function filterCase2() {
        if (!state.reconciledData) return;
        
        const q = case2Search.value.toLowerCase().trim();
        const selectedCarrier = case2CarrierFilter.value;
        
        const filtered = state.reconciledData.case2.filter(item => {
            const matchesQuery = !q || 
                item.order_id.toLowerCase().includes(q) ||
                (item.recipient && item.recipient.toLowerCase().includes(q)) ||
                (item.phone && item.phone.toLowerCase().includes(q)) ||
                (item.tracking_id && item.tracking_id.toLowerCase().includes(q)) ||
                (item.ivy_code && item.ivy_code.toLowerCase().includes(q));
                
            const matchesCarrier = !selectedCarrier || item.carrier === selectedCarrier;
            
            return matchesQuery && matchesCarrier;
        });
        
        renderCase2Table(filtered);
    }

    function renderCase2Table(data) {
        case2Tbody.innerHTML = '';
        
        if (data.length === 0) {
            case2Tbody.innerHTML = '<tr><td colspan="10" class="text-center">Không tìm thấy đơn hàng lệch khớp nào.</td></tr>';
            case2PaginationInfo.textContent = 'Đang hiển thị 0/0 đơn hàng bị lệch.';
            return;
        }

        data.forEach((item, idx) => {
            const tr = document.createElement('tr');
            
            tr.innerHTML = `
                <td>${idx + 1}</td>
                <td class="font-semibold text-primary">${item.order_id}</td>
                <td class="text-indigo">${item.ivy_code || '-'}</td>
                <td>${item.recipient}</td>
                <td>${item.phone}</td>
                <td>${item.carrier}</td>
                <td>${item.tracking_id || '-'}</td>
                <td style="max-width: 200px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;" title="${item.cancel_reason}">${item.cancel_reason}</td>
                <td><span class="badge-status badge-status-danger">${item.platform_status}</span></td>
                <td><span class="badge-status badge-status-warning">${item.internal_status}</span></td>
            `;
            case2Tbody.appendChild(tr);
        });

        case2PaginationInfo.textContent = `Đang hiển thị ${data.length}/${state.reconciledData.case2.length} đơn hàng bị lệch.`;
    }

    // 9. Excel/CSV Export utilities
    exportCase1Btn.addEventListener('click', () => {
        if (!state.reconciledData || state.reconciledData.case1.length === 0) return;
        
        const headers = ["STT", "Mã Đơn Hàng (Sàn)", "Mã Sản Phẩm", "Mã IVY (Kho)", "Người Mua", "Lý Do Trả Hàng", "Trạng Thái TikTok", "Trạng Thái Kho", "Thời Gian Hoàn Tiền", "Mã Vận Đơn Trả"];
        const rows = state.reconciledData.case1.map((item, idx) => [
            idx + 1,
            item.order_id,
            item.product_code,
            item.ivy_code,
            item.buyer,
            item.reason,
            "Completed Return",
            item.internal_status,
            item.refund_time,
            item.tracking_id
        ]);
        
        downloadCSV(headers, rows, "Case1_TraHang_ChuaNhapKho.csv");
    });

    exportCase2Btn.addEventListener('click', () => {
        if (!state.reconciledData || state.reconciledData.case2.length === 0) return;
        
        const headers = ["STT", "Mã Đơn Hàng (Sàn)", "Mã IVY (Kho)", "Người Nhận", "Số Điện Thoại", "Đơn Vị Vận Chuyển", "Mã Vận Đơn", "Lý Do Hủy", "Trạng Thái TikTok", "Trạng Thái Kho"];
        const rows = state.reconciledData.case2.map((item, idx) => [
            idx + 1,
            item.order_id,
            item.ivy_code,
            item.recipient,
            item.phone,
            item.carrier,
            item.tracking_id,
            item.cancel_reason,
            item.platform_status,
            item.internal_status
        ]);
        
        downloadCSV(headers, rows, "Case2_DonHuy_ChuaHoanKho.csv");
    });

    function downloadCSV(headers, rows, filename) {
        const formatCSVCell = cell => {
            if (cell === null || cell === undefined) return '';
            const str = String(cell).replace(/"/g, '""');
            return `"${str}"`;
        };

        let csvContent = '\uFEFF'; // UTF-8 BOM
        csvContent += headers.map(formatCSVCell).join(',') + '\n';
        
        rows.forEach(row => {
            csvContent += row.map(formatCSVCell).join(',') + '\n';
        });
        
        const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
        const url = URL.createObjectURL(blob);
        const link = document.createElement("a");
        link.setAttribute("href", url);
        link.setAttribute("download", filename);
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
    }

    // 10. Modal Close utilities
    closeModalBtns.forEach(btn => {
        btn.addEventListener('click', () => {
            errorModal.classList.add('hidden');
        });
    });

    window.addEventListener('click', (e) => {
        if (e.target === errorModal) {
            errorModal.classList.add('hidden');
        }
    });

    // Helper functions
    function showError(message) {
        errorModalMessage.innerHTML = message;
        errorModal.classList.remove('hidden');
    }

    function formatNumber(num) {
        return num.toString().replace(/\B(?=(\d{3})+(?!\d))/g, ".");
    }

    function formatBytes(bytes) {
        if (bytes === 0) return '0 Bytes';
        const k = 1024;
        const sizes = ['Bytes', 'KB', 'MB', 'GB'];
        const i = Math.floor(Math.log(bytes) / Math.log(k));
        return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
    }
});
