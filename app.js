const STORAGE_KEY = "ftdWebReservations";
const API_BASE = "";
const WEEKDAY = ["周日", "周一", "周二", "周三", "周四", "周五", "周六"];
const roles = ["教员", "学员"];

let viewMode = "future";
let reservations = [];
let selectedDetailId = "";

const els = {
  visibleCount: document.querySelector("#visibleCount"),
  futureTab: document.querySelector("#futureTab"),
  historyTab: document.querySelector("#historyTab"),
  rangeText: document.querySelector("#rangeText"),
  maxDateText: document.querySelector("#maxDateText"),
  schedule: document.querySelector("#schedule"),
  newBooking: document.querySelector("#newBooking"),
  exportExcel: document.querySelector("#exportExcel"),
  bookingDialog: document.querySelector("#bookingDialog"),
  bookingForm: document.querySelector("#bookingForm"),
  closeDialog: document.querySelector("#closeDialog"),
  cancelBooking: document.querySelector("#cancelBooking"),
  bookingDate: document.querySelector("#bookingDate"),
  startTime: document.querySelector("#startTime"),
  endTime: document.querySelector("#endTime"),
  addPerson: document.querySelector("#addPerson"),
  peopleList: document.querySelector("#peopleList"),
  detailDialog: document.querySelector("#detailDialog"),
  detailDate: document.querySelector("#detailDate"),
  detailTime: document.querySelector("#detailTime"),
  detailPeople: document.querySelector("#detailPeople"),
  closeDetail: document.querySelector("#closeDetail"),
  okDetail: document.querySelector("#okDetail"),
  deleteBooking: document.querySelector("#deleteBooking")
};

function pad(value) {
  return value < 10 ? `0${value}` : `${value}`;
}

function dateKey(date) {
  return `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())}`;
}

function parseDate(key) {
  const parts = key.split("-").map(Number);
  return new Date(parts[0], parts[1] - 1, parts[2]);
}

function addDays(date, days) {
  const next = new Date(date);
  next.setDate(next.getDate() + days);
  return next;
}

function addMonths(date, months) {
  const next = new Date(date);
  next.setMonth(next.getMonth() + months);
  return next;
}

function formatDateLabel(key) {
  const date = parseDate(key);
  return `${pad(date.getMonth() + 1)}月${pad(date.getDate())}日`;
}

function formatWeekday(key) {
  return WEEKDAY[parseDate(key).getDay()];
}

function toMinutes(time) {
  const [hour, minute] = time.split(":").map(Number);
  return hour * 60 + minute;
}

function isBookableDate(date) {
  const day = date.getDay();
  return day >= 3 && day <= 5;
}

async function loadReservations() {
  try {
    const response = await fetch(`${API_BASE}/api/reservations`, { cache: "no-store" });
    if (!response.ok) throw new Error("api failed");
    return await response.json();
  } catch (error) {
    if (!isLocalPreview()) {
      alert("在线预约记录读取失败，请稍后刷新页面。");
      return [];
    }
    return loadLocalReservations();
  }
}

function isLocalPreview() {
  return ["", "localhost", "127.0.0.1"].includes(location.hostname);
}

function loadLocalReservations() {
  try {
    return JSON.parse(localStorage.getItem(STORAGE_KEY) || "[]");
  } catch (error) {
    return [];
  }
}

async function saveReservations(next) {
  try {
    const response = await fetch(`${API_BASE}/api/reservations`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(next)
    });
    if (!response.ok) throw new Error("save failed");
    reservations = await response.json();
    localStorage.setItem(STORAGE_KEY, JSON.stringify(reservations));
  } catch (error) {
    if (!isLocalPreview()) {
      throw new Error("在线保存失败，请稍后再试。");
    }
    reservations = next;
    localStorage.setItem(STORAGE_KEY, JSON.stringify(next));
    throw new Error("在线保存失败，请检查服务器是否启动。");
  }
}

function getRoleNames(people, role) {
  const names = people.filter((item) => item.role === role).map((item) => item.name);
  return names.length ? names.join("、") : "无";
}

function normalizeReservation(item) {
  return {
    ...item,
    timeText: `${item.start}-${item.end}`,
    teachersText: getRoleNames(item.people || [], "教员"),
    studentsText: getRoleNames(item.people || [], "学员")
  };
}

function formatTimeLabel(start, end) {
  return `${start.replace(":", "")}-${end.replace(":", "")}`;
}

function rangesOverlap(leftStart, leftEnd, rightStart, rightEnd) {
  return Math.max(toMinutes(leftStart), toMinutes(rightStart)) < Math.min(toMinutes(leftEnd), toMinutes(rightEnd));
}

function findConflict(date, start, end) {
  return reservations.find((item) => item.date === date && rangesOverlap(start, end, item.start, item.end));
}

function escapeHtml(value) {
  return String(value == null ? "" : value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function escapeXml(value) {
  return escapeHtml(value);
}

function getRange() {
  const today = new Date();
  const todayKey = dateKey(today);
  const maxDate = addMonths(today, 1);
  return { startDate: parseDate(todayKey), endDate: maxDate, todayKey, maxDate };
}

function buildDayColumns(dayReservations) {
  const seen = {};
  dayReservations.forEach((item) => {
    const key = `${item.start}-${item.end}`;
    seen[key] = {
      label: formatTimeLabel(item.start, item.end),
      start: item.start,
      end: item.end
    };
  });

  const columns = Object.keys(seen).map((key) => seen[key]).sort((a, b) => {
    const startDiff = toMinutes(a.start) - toMinutes(b.start);
    return startDiff || toMinutes(a.end) - toMinutes(b.end);
  });

  return columns.length ? columns : [{ label: "预约时间", start: "08:00", end: "10:00", empty: true }];
}

function buildCells(dayReservations, columns) {
  return columns.map((column) => {
    const slotBookings = dayReservations.filter((item) => item.start === column.start && item.end === column.end);
    return {
      ...column,
      hasBooking: slotBookings.length > 0,
      reservationId: slotBookings[0]?.id || "",
      text: slotBookings.map((booking) => `教员：${booking.teachersText}\n学员：${booking.studentsText}`).join("\n\n")
    };
  });
}

function buildRows(startDate, endDate) {
  const dayCount = Math.floor((endDate - startDate) / 86400000) + 1;
  return Array.from({ length: dayCount }).map((_, index) => {
    const current = addDays(startDate, index);
    const key = dateKey(current);
    const dayReservations = reservations
      .filter((item) => item.date === key)
      .sort((a, b) => toMinutes(a.start) - toMinutes(b.start))
      .map(normalizeReservation);
    const columns = buildDayColumns(dayReservations);
    return {
      date: key,
      label: formatDateLabel(key),
      weekday: formatWeekday(key),
      reservations: dayReservations,
      cells: buildCells(dayReservations, columns),
      hasReservations: dayReservations.length > 0
    };
  }).filter((row) => isBookableDate(parseDate(row.date)));
}

function buildReservationRows() {
  const dates = [...new Set(reservations.map((item) => item.date))]
    .filter((key) => /^\d{4}-\d{2}-\d{2}$/.test(key))
    .sort((a, b) => parseDate(a) - parseDate(b));

  return dates.map((key) => {
    const dayReservations = reservations
      .filter((item) => item.date === key)
      .sort((a, b) => toMinutes(a.start) - toMinutes(b.start))
      .map(normalizeReservation);
    const columns = buildDayColumns(dayReservations);
    return {
      date: key,
      label: formatDateLabel(key),
      weekday: formatWeekday(key),
      reservations: dayReservations,
      cells: buildCells(dayReservations, columns),
      hasReservations: dayReservations.length > 0
    };
  });
}

function getVisibleRows(startDate, endDate) {
  return viewMode === "history" ? buildReservationRows() : buildRows(startDate, endDate);
}

function formatVisibleRange(rows, startDate, endDate) {
  if (!rows.length) return "暂无预约记录";
  const first = rows[0].date;
  const last = rows[rows.length - 1].date;
  return viewMode === "history"
    ? `${formatDateLabel(first)} - ${formatDateLabel(last)}`
    : `${formatDateLabel(dateKey(startDate))} - ${formatDateLabel(dateKey(endDate))}`;
}

function renderSchedule() {
  const { startDate, endDate, todayKey, maxDate } = getRange();
  const rows = getVisibleRows(startDate, endDate);
  const visibleCount = rows.reduce((total, row) => total + row.reservations.length, 0);

  els.visibleCount.textContent = visibleCount;
  els.maxDateText.textContent = formatDateLabel(dateKey(maxDate));
  els.rangeText.textContent = formatVisibleRange(rows, startDate, endDate);
  els.futureTab.classList.toggle("active", viewMode === "future");
  els.historyTab.classList.toggle("active", viewMode === "history");

  els.schedule.innerHTML = [
    `<div class="row head-row"><div class="cell date-cell">日期</div><div class="cell day-head">预约时间 / 使用人员</div></div>`,
    ...rows.map((row) => renderRow(row, todayKey, maxDate))
  ].join("");

  els.schedule.querySelectorAll("[data-add-date]").forEach((node) => {
    node.addEventListener("click", () => openBooking(node.dataset.addDate, node.dataset.start || "", node.dataset.end || ""));
  });
  els.schedule.querySelectorAll("[data-detail-id]").forEach((node) => {
    node.addEventListener("click", (event) => {
      event.stopPropagation();
      openDetail(node.dataset.detailId);
    });
  });
}

function renderRow(row, todayKey, maxDate) {
  const rowDate = parseDate(row.date);
  const bookable = rowDate >= parseDate(todayKey) && rowDate <= maxDate && isBookableDate(rowDate);
  const cells = row.cells.map((cell) => {
    const classes = `cell slot-cell ${cell.hasBooking ? "booked" : ""}`;
    if (cell.hasBooking) {
      return `<div class="${classes}" data-detail-id="${escapeHtml(cell.reservationId)}">
        <div class="slot-time">${escapeHtml(cell.label)}</div>
        <div>${escapeHtml(cell.text)}</div>
      </div>`;
    }
    return `<div class="${classes} empty-cell" data-add-date="${escapeHtml(row.date)}" data-start="${escapeHtml(cell.start)}" data-end="${escapeHtml(cell.end)}">
      <div class="slot-time">${escapeHtml(cell.label)}</div>
      <div>${bookable ? "教员：\n\n学员：" : "不可预约"}</div>
    </div>`;
  }).join("");

  return `<div class="row">
    <div class="cell date-cell" data-add-date="${escapeHtml(row.date)}">${escapeHtml(row.label)}<br>${escapeHtml(row.weekday)}</div>
    ${cells}
  </div>`;
}

function getDefaultBookingDate() {
  const { todayKey, maxDate } = getRange();
  let current = parseDate(todayKey);
  while (current <= maxDate) {
    if (isBookableDate(current)) return dateKey(current);
    current = addDays(current, 1);
  }
  return todayKey;
}

function blankPerson(role = "学员") {
  return { name: "", employeeNo: "", role };
}

function renderPeople(people) {
  els.peopleList.innerHTML = people.map((person, index) => `
    <div class="person-row" data-person-row="${index}">
      <label>姓名<input value="${escapeHtml(person.name)}" data-field="name"></label>
      <label>工号<input value="${escapeHtml(person.employeeNo)}" data-field="employeeNo"></label>
      <label>身份
        <select data-field="role">
          ${roles.map((role) => `<option value="${role}" ${person.role === role ? "selected" : ""}>${role}</option>`).join("")}
        </select>
      </label>
      <button class="remove-person" type="button" ${people.length <= 1 ? "disabled" : ""}>删除</button>
    </div>
  `).join("");
}

function readPeople() {
  return Array.from(els.peopleList.querySelectorAll("[data-person-row]")).map((row) => ({
    name: row.querySelector('[data-field="name"]').value.trim(),
    employeeNo: row.querySelector('[data-field="employeeNo"]').value.trim(),
    role: row.querySelector('[data-field="role"]').value
  }));
}

function openBooking(date, start, end) {
  const { todayKey, maxDate } = getRange();
  const target = date || getDefaultBookingDate();
  if (parseDate(target) < parseDate(todayKey)) {
    alert("历史记录不可新增预约");
    return;
  }
  if (parseDate(target) > maxDate) {
    alert("超过一个月预约范围");
    return;
  }
  if (!isBookableDate(parseDate(target))) {
    alert("仅周三至周五可预约");
    return;
  }

  els.bookingDate.value = target;
  els.bookingDate.min = todayKey;
  els.bookingDate.max = dateKey(maxDate);
  els.startTime.value = start || "08:00";
  els.endTime.value = end || "10:00";
  renderPeople([blankPerson()]);
  els.bookingDialog.showModal();
}

function openDetail(id) {
  const detail = reservations.find((item) => item.id === id);
  if (!detail) return;
  selectedDetailId = id;
  const booking = normalizeReservation(detail);
  els.detailDate.textContent = `${booking.date} ${formatWeekday(booking.date)}`;
  els.detailTime.textContent = booking.timeText;
  els.detailPeople.innerHTML = booking.people.map((person) => `
    <div class="detail-person"><strong>${escapeHtml(person.name)} · ${escapeHtml(person.role)}</strong><br><span class="muted">工号：${escapeHtml(person.employeeNo)}</span></div>
  `).join("");
  els.detailDialog.showModal();
}

async function submitBooking(event) {
  event.preventDefault();
  const date = els.bookingDate.value;
  const start = els.startTime.value;
  const end = els.endTime.value;
  const { todayKey, maxDate } = getRange();

  if (!date || !start || !end) {
    alert("请填写预约日期和时间");
    return;
  }
  if (parseDate(date) < parseDate(todayKey) || parseDate(date) > maxDate) {
    alert("只能提前一个月预约");
    return;
  }
  if (!isBookableDate(parseDate(date))) {
    alert("仅周三至周五可预约");
    return;
  }
  if (toMinutes(end) <= toMinutes(start)) {
    alert("结束时间需晚于开始时间");
    return;
  }

  const people = readPeople();
  if (people.some((person) => !person.name || !person.employeeNo || !person.role)) {
    alert("请补全姓名、工号和身份");
    return;
  }
  if (people.length > 5) {
    alert("每场最多5人");
    return;
  }

  const conflict = findConflict(date, start, end);
  if (conflict) {
    alert(`${conflict.start}-${conflict.end} 已有预约，请选择其他时间。`);
    return;
  }

  try {
    await saveReservations([...reservations, {
      id: `${Date.now()}-${Math.random().toString(16).slice(2)}`,
      date,
      start,
      end,
      people,
      createdAt: new Date().toISOString()
    }]);
    els.bookingDialog.close();
    renderSchedule();
  } catch (error) {
    alert(error.message);
  }
}

function excelCell(value, styleId) {
  return `<Cell ss:StyleID="${styleId}"><Data ss:Type="String">${escapeXml(value)}</Data></Cell>`;
}

function reservationExcelText(reservation) {
  const teachers = reservation.people.filter((person) => person.role === "教员").map((person) => person.name).join("、") || "无";
  const students = reservation.people.filter((person) => person.role === "学员").map((person) => person.name);
  const firstStudent = students[0] || "";
  const otherStudents = students.slice(1).join("、");
  return `${reservation.timeText.replace(/:/g, "")}\n教员：${teachers}\n学员：${firstStudent}\n学员：${otherStudents}`;
}

function buildExcelXml(rows, rangeText) {
  const maxCells = Math.max(1, ...rows.map((day) => Math.max(1, day.reservations.length)));
  const tableRows = rows.map((day) => {
    const cells = day.reservations.map((reservation) => excelCell(reservationExcelText(reservation), "Wrap"));
    while (cells.length < maxCells) cells.push(excelCell("", "Blank"));
    return `<Row ss:Height="150">${excelCell(`${formatDateLabel(day.date)}\n${day.weekday}`, "Date")}${cells.join("")}</Row>`;
  }).join("\n   ");

  return `<?xml version="1.0" encoding="UTF-8"?>
<?mso-application progid="Excel.Sheet"?>
<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"
 xmlns:o="urn:schemas-microsoft-com:office:office"
 xmlns:x="urn:schemas-microsoft-com:office:excel"
 xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"
 xmlns:html="http://www.w3.org/TR/REC-html40">
 <DocumentProperties xmlns="urn:schemas-microsoft-com:office:office">
  <Title>${escapeXml(`飞行部FTD使用记录 ${rangeText}`)}</Title>
 </DocumentProperties>
 <Styles>
  <Style ss:ID="Date"><Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/><Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#D4DEDC"/><Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#D4DEDC"/><Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#D4DEDC"/><Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#D4DEDC"/></Borders><Font ss:FontName="Arial" ss:Size="11" ss:Bold="1" ss:Color="#172423"/></Style>
  <Style ss:ID="Wrap"><Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/><Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#E0E5E3"/><Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#E0E5E3"/><Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#E0E5E3"/><Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#E0E5E3"/></Borders><Font ss:FontName="Arial" ss:Size="11" ss:Bold="1" ss:Color="#172423"/><Interior ss:Color="#FBF1E2" ss:Pattern="Solid"/></Style>
  <Style ss:ID="Blank"><Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/><Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#E0E5E3"/><Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#E0E5E3"/><Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#E0E5E3"/><Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#E0E5E3"/></Borders><Font ss:FontName="Arial" ss:Size="11" ss:Bold="1" ss:Color="#172423"/></Style>
 </Styles>
 <Worksheet ss:Name="FTD使用记录">
  <Table>
   <Column ss:Width="136"/>
   ${Array.from({ length: maxCells }).map(() => '<Column ss:Width="228"/>').join("\n   ")}
   ${tableRows}
  </Table>
  <WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel">
   <PageSetup><Layout x:Orientation="Landscape"/></PageSetup>
   <FitToPage/>
  </WorksheetOptions>
 </Worksheet>
</Workbook>`;
}

function exportExcel() {
  const { startDate, endDate } = getRange();
  const rows = getVisibleRows(startDate, endDate);
  const rangeText = viewMode === "history" && rows.length
    ? `${formatDateLabel(rows[0].date)}-${formatDateLabel(rows[rows.length - 1].date)}`
    : `${formatDateLabel(dateKey(startDate))}-${formatDateLabel(dateKey(endDate))}`;
  const xml = buildExcelXml(rows, rangeText);
  const blob = new Blob([xml], { type: "application/vnd.ms-excel;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = `飞行部FTD使用记录_${rangeText}.xls`;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}

els.futureTab.addEventListener("click", () => {
  viewMode = "future";
  renderSchedule();
});

els.historyTab.addEventListener("click", async () => {
  viewMode = "history";
  reservations = await loadReservations();
  renderSchedule();
});

els.newBooking.addEventListener("click", () => openBooking());
els.exportExcel.addEventListener("click", exportExcel);
els.closeDialog.addEventListener("click", () => els.bookingDialog.close());
els.cancelBooking.addEventListener("click", () => els.bookingDialog.close());
els.bookingForm.addEventListener("submit", submitBooking);
els.addPerson.addEventListener("click", () => {
  const people = readPeople();
  if (people.length >= 5) {
    alert("最多5人");
    return;
  }
  people.push(blankPerson());
  renderPeople(people);
});
els.peopleList.addEventListener("click", (event) => {
  if (!event.target.classList.contains("remove-person")) return;
  const people = readPeople();
  const row = event.target.closest("[data-person-row]");
  people.splice(Number(row.dataset.personRow), 1);
  renderPeople(people);
});
els.closeDetail.addEventListener("click", () => els.detailDialog.close());
els.okDetail.addEventListener("click", () => els.detailDialog.close());
els.deleteBooking.addEventListener("click", async () => {
  if (!selectedDetailId) return;
  if (!confirm("确认取消这场FTD预约吗？")) return;
  try {
    await saveReservations(reservations.filter((item) => item.id !== selectedDetailId));
    selectedDetailId = "";
    els.detailDialog.close();
    renderSchedule();
  } catch (error) {
    alert(error.message);
  }
});

async function init() {
  reservations = await loadReservations();
  renderSchedule();
}

init();
