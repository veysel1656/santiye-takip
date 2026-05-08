import { useEffect, useMemo, useState } from "react";
import { supabase } from "../supabase";
import dynamic from "next/dynamic";
import * as XLSX from "xlsx";
import jsPDF from "jspdf";
import autoTable from "jspdf-autotable";

const Chart = dynamic(() => import("react-apexcharts"), { ssr: false });

export default function Home() {
  const [session, setSession] = useState(null);
  const [userRole, setUserRole] = useState("office");
  const [selectedPier, setSelectedPier] = useState("P1");

  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [loginError, setLoginError] = useState("");

  const [productions, setProductions] = useState([]);
  const [steel, setSteel] = useState([]);
  const [logs, setLogs] = useState([]);
  const [progress, setProgress] = useState([]);

  const [productionName, setProductionName] = useState("");
  const [productionQty, setProductionQty] = useState("");

  const [steelName, setSteelName] = useState("");
  const [steelQty, setSteelQty] = useState("");
  const [steelType, setSteelType] = useState("giriş");

  const todayStr = new Date().toISOString().split("T")[0];

  const [reportDate, setReportDate] = useState(todayStr);
  const [weather, setWeather] = useState("");
  const [team, setTeam] = useState("");
  const [logText, setLogText] = useState("");
  const [issue, setIssue] = useState("");
  const [reportImage, setReportImage] = useState(null);
  const [isSavingLog, setIsSavingLog] = useState(false);

  const [editingLogId, setEditingLogId] = useState(null);
  const [existingImageUrl, setExistingImageUrl] = useState("");

  const [progressItem, setProgressItem] = useState("");
  const [progressTotal, setProgressTotal] = useState("");
  const [progressDone, setProgressDone] = useState("");

  const [filterStartDate, setFilterStartDate] = useState("");
  const [filterEndDate, setFilterEndDate] = useState("");
  const [filterText, setFilterText] = useState("");
  const [onlyWithImage, setOnlyWithImage] = useState(false);

  const [lastRefresh, setLastRefresh] = useState("");
  const [selectedImage, setSelectedImage] = useState("");

  const stockThresholds = {
    "Ø8": 5,
    "Ø10": 5,
    "Ø12": 8,
    "Ø14": 8,
    "Ø16": 10,
    "Ø18": 10,
    "Ø20": 12,
    "Ø22": 12,
    "Ø25": 15,
    "Nervürlü Demir": 20,
  };

  async function loadRole(userEmail) {
    const { data, error } = await supabase
      .from("users")
      .select("role, email")
      .eq("email", userEmail)
      .maybeSingle();

    if (error) {
      console.error("ROLE LOAD ERROR:", error);
      setUserRole("office");
      return;
    }

    const role = data?.role === "chief" ? "chief" : "office";
    setUserRole(role);
  }

  async function loadData() {
    const { data: p1 } = await supabase
      .from("productions")
      .select("*")
      .order("created_at", { ascending: false });

    const { data: p2 } = await supabase
      .from("steel_stock")
      .select("*")
      .order("created_at", { ascending: false });

    const { data: p3 } = await supabase
      .from("daily_logs")
      .select("*")
      .order("report_date", { ascending: false })
      .order("created_at", { ascending: false });

    const { data: p4 } = await supabase
      .from("progress")
      .select("*")
      .order("created_at", { ascending: false });

    setProductions(p1 || []);
    setSteel(p2 || []);
    setLogs(p3 || []);
    setProgress(p4 || []);
    setLastRefresh(new Date().toLocaleTimeString("tr-TR"));
  }

  useEffect(() => {
    supabase.auth.getSession().then(async ({ data }) => {
      const currentSession = data.session;
      setSession(currentSession || null);

      if (currentSession?.user?.email) {
        await loadRole(currentSession.user.email);
        await loadData();
      }
    });

    const {
      data: { subscription },
    } = supabase.auth.onAuthStateChange(async (_event, newSession) => {
      setSession(newSession || null);

      if (newSession?.user?.email) {
        await loadRole(newSession.user.email);
        await loadData();
      } else {
        setUserRole("office");
      }
    });

    return () => subscription.unsubscribe();
  }, []);

  useEffect(() => {
    if (!session) return;

    const interval = setInterval(() => {
      loadData();
    }, 15000);

    return () => clearInterval(interval);
  }, [session]);

  async function signIn() {
    setLoginError("");

    const { error } = await supabase.auth.signInWithPassword({
      email,
      password,
    });

    if (error) {
      setLoginError(error.message || "Giriş yapılamadı");
    }
  }

  async function signOut() {
    await supabase.auth.signOut();
    setSession(null);
    setUserRole("office");
  }

  const canEdit = userRole === "chief";

  async function addProduction() {
    if (!canEdit) {
      alert("Bu işlem için yetkin yok");
      return;
    }

    if (!productionName.trim()) {
      alert("İmalat adı boş olamaz");
      return;
    }

    if (!productionQty || Number(productionQty) <= 0) {
      alert("Geçerli miktar gir");
      return;
    }

    const { error } = await supabase.from("productions").insert([
      {
        name: productionName.trim(),
        quantity: Number(productionQty),
        pier: selectedPier,
      },
    ]);

    if (error) {
      alert("İmalat kayıt hatası: " + error.message);
      return;
    }

    setProductionName("");
    setProductionQty("");
    loadData();
  }

  async function addSteel() {
    if (!canEdit) {
      alert("Bu işlem için yetkin yok");
      return;
    }

    if (!steelName.trim()) {
      alert("Malzeme adı boş olamaz");
      return;
    }

    if (!steelQty || Number(steelQty) <= 0) {
      alert("Geçerli miktar gir");
      return;
    }

    const { error } = await supabase.from("steel_stock").insert([
      {
        name: steelName.trim(),
        quantity: Number(steelQty),
        type: steelType,
        pier: selectedPier,
      },
    ]);

    if (error) {
      alert("Stok kayıt hatası: " + error.message);
      return;
    }

    setSteelName("");
    setSteelQty("");
    setSteelType("giriş");
    loadData();
  }

  async function addProgress() {
    if (!canEdit) {
      alert("Bu işlem için yetkin yok");
      return;
    }

    if (!progressItem.trim()) {
      alert("İş kalemi boş olamaz");
      return;
    }

    if (!progressTotal || Number(progressTotal) <= 0) {
      alert("Toplam miktar geçerli olmalı");
      return;
    }

    if (progressDone === "" || Number(progressDone) < 0) {
      alert("Tamamlanan miktar geçerli olmalı");
      return;
    }

    const { error } = await supabase.from("progress").insert([
      {
        item: progressItem.trim(),
        total_quantity: Number(progressTotal),
        completed_quantity: Number(progressDone),
        pier: selectedPier,
      },
    ]);

    if (error) {
      alert("Hakediş kayıt hatası: " + error.message);
      return;
    }

    setProgressItem("");
    setProgressTotal("");
    setProgressDone("");
    loadData();
  }

  async function uploadReportImage(file) {
    if (!file) return "";

    const fileExt = file.name.split(".").pop();
    const fileName = `${Date.now()}-${Math.random()
      .toString(36)
      .slice(2)}.${fileExt}`;

    const { error: uploadError } = await supabase.storage
      .from("daily-report-images")
      .upload(fileName, file, {
        cacheControl: "3600",
        upsert: false,
      });

    if (uploadError) {
      console.error("UPLOAD ERROR:", uploadError);
      alert("Fotoğraf yükleme hatası: " + uploadError.message);
      return "";
    }

    const { data } = supabase.storage
      .from("daily-report-images")
      .getPublicUrl(fileName);

    return data?.publicUrl || "";
  }

  function resetLogForm() {
    setEditingLogId(null);
    setExistingImageUrl("");
    setReportDate(todayStr);
    setWeather("");
    setTeam("");
    setLogText("");
    setIssue("");
    setReportImage(null);
  }

  async function addLog() {
    if (!canEdit) {
      alert("Bu işlem için yetkin yok");
      return;
    }

    if (isSavingLog) return;

    if (!reportDate) {
      alert("Tarih seçmelisin");
      return;
    }

    if (!logText.trim()) {
      alert("Yapılan iş alanı boş olamaz");
      return;
    }

    setIsSavingLog(true);

    try {
      let imageUrl = existingImageUrl || "";

      if (reportImage) {
        imageUrl = await uploadReportImage(reportImage);
        if (!imageUrl) {
          setIsSavingLog(false);
          return;
        }
      }

      const payload = {
        pier: selectedPier,
        report_date: reportDate,
        weather: weather.trim(),
        team: team.trim(),
        description: logText.trim(),
        issue: issue.trim(),
        image_url: imageUrl,
      };

      let error = null;

      if (editingLogId) {
        const result = await supabase
          .from("daily_logs")
          .update(payload)
          .eq("id", editingLogId);
        error = result.error;
      } else {
        const result = await supabase.from("daily_logs").insert([payload]);
        error = result.error;
      }

      if (error) {
        console.error("SAVE ERROR:", error);
        alert("Rapor kayıt hatası: " + error.message);
        setIsSavingLog(false);
        return;
      }

      alert(editingLogId ? "Rapor güncellendi" : "Rapor kaydedildi");
      resetLogForm();
      await loadData();
    } catch (err) {
      console.error("SAVE ERROR:", err);
      alert("Beklenmeyen bir hata oluştu");
    } finally {
      setIsSavingLog(false);
    }
  }

  function startEditLog(item) {
    if (!canEdit) return;

    setEditingLogId(item.id);
    setReportDate(item.report_date || todayStr);
    setWeather(item.weather || "");
    setTeam(item.team || "");
    setLogText(item.description || "");
    setIssue(item.issue || "");
    setExistingImageUrl(item.image_url || "");
    setReportImage(null);
    window.scrollTo({ top: 0, behavior: "smooth" });
  }

  async function deleteLog(id) {
    if (!canEdit) {
      alert("Bu işlem için yetkin yok");
      return;
    }

    const ok = window.confirm("Bu raporu silmek istediğine emin misin?");
    if (!ok) return;

    const { error } = await supabase.from("daily_logs").delete().eq("id", id);

    if (error) {
      alert("Silme hatası: " + error.message);
      return;
    }

    if (editingLogId === id) {
      resetLogForm();
    }

    await loadData();
    alert("Rapor silindi");
  }

  const filteredProductions = useMemo(
    () => productions.filter((item) => (item.pier || "P1") === selectedPier),
    [productions, selectedPier]
  );

  const filteredSteel = useMemo(
    () => steel.filter((item) => (item.pier || "P1") === selectedPier),
    [steel, selectedPier]
  );

  const filteredLogs = useMemo(() => {
    return logs.filter((item) => {
      if ((item.pier || "P1") !== selectedPier) return false;
      if (filterStartDate && item.report_date && item.report_date < filterStartDate)
        return false;
      if (filterEndDate && item.report_date && item.report_date > filterEndDate)
        return false;
      if (onlyWithImage && !item.image_url) return false;

      if (filterText) {
        const q = filterText.toLowerCase();
        const inDescription = item.description?.toLowerCase().includes(q);
        const inTeam = item.team?.toLowerCase().includes(q);
        const inWeather = item.weather?.toLowerCase().includes(q);
        const inIssue = item.issue?.toLowerCase().includes(q);
        if (!(inDescription || inTeam || inWeather || inIssue)) return false;
      }

      return true;
    });
  }, [
    logs,
    selectedPier,
    filterStartDate,
    filterEndDate,
    filterText,
    onlyWithImage,
  ]);

  const filteredProgress = useMemo(
    () => progress.filter((item) => (item.pier || "P1") === selectedPier),
    [progress, selectedPier]
  );

  const totalProduction = filteredProductions.reduce(
    (sum, item) => sum + Number(item.quantity || 0),
    0
  );

  const totalSteel = filteredSteel.reduce((sum, item) => {
    const qty = Number(item.quantity || 0);
    return item.type === "çıkış" ? sum - qty : sum + qty;
  }, 0);

  const avgProgress =
    filteredProgress.length > 0
      ? (
          filteredProgress.reduce((sum, item) => {
            const total = Number(item.total_quantity || 0);
            const done = Number(item.completed_quantity || 0);
            if (!total) return sum;
            return sum + (done / total) * 100;
          }, 0) / filteredProgress.length
        ).toFixed(1)
      : "0";

  const pierTotals = ["P1", "P2", "P3", "P4"].map((pier) =>
    productions
      .filter((item) => (item.pier || "P1") === pier)
      .reduce((sum, item) => sum + Number(item.quantity || 0), 0)
  );

  const maxPierData = useMemo(() => {
    const piers = ["P1", "P2", "P3", "P4"].map((pier) => ({
      pier,
      total: productions
        .filter((item) => (item.pier || "P1") === pier)
        .reduce((sum, item) => sum + Number(item.quantity || 0), 0),
    }));

    return piers.sort((a, b) => b.total - a.total)[0] || {
      pier: "-",
      total: 0,
    };
  }, [productions]);

  const recentLogs = useMemo(() => filteredLogs.slice(0, 5), [filteredLogs]);

  const todayProductionTotal = useMemo(() => {
    return filteredProductions
      .filter(
        (item) => item.created_at && item.created_at.slice(0, 10) === todayStr
      )
      .reduce((sum, item) => sum + Number(item.quantity || 0), 0);
  }, [filteredProductions, todayStr]);

  const todayLogs = useMemo(() => {
    return filteredLogs.filter((item) => {
      if (item.report_date) return item.report_date === todayStr;
      if (item.created_at) return item.created_at.slice(0, 10) === todayStr;
      return false;
    });
  }, [filteredLogs, todayStr]);

  const thisWeekLogsCount = useMemo(() => {
    const now = new Date();
    const start = new Date(now);
    const day = now.getDay() || 7;
    start.setDate(now.getDate() - day + 1);
    start.setHours(0, 0, 0, 0);

    return filteredLogs.filter((item) => {
      if (!item.report_date) return false;
      return new Date(item.report_date) >= start;
    }).length;
  }, [filteredLogs]);

  const todayWorkSummary = useMemo(() => {
    return todayLogs
      .map((item) => item.description)
      .filter(Boolean)
      .slice(0, 3);
  }, [todayLogs]);

  const productionByDay = useMemo(() => {
    const grouped = {};

    filteredProductions.forEach((item) => {
      const date = item.created_at
        ? new Date(item.created_at).toLocaleDateString("tr-TR")
        : "Tarihsiz";
      grouped[date] = (grouped[date] || 0) + Number(item.quantity || 0);
    });

    const entries = Object.entries(grouped).reverse().slice(-7);
    return {
      categories: entries.map(([date]) => date),
      values: entries.map(([, value]) => value),
    };
  }, [filteredProductions]);

  const logsByDay = useMemo(() => {
    const grouped = {};

    filteredLogs.forEach((item) => {
      const date = item.report_date
        ? new Date(item.report_date).toLocaleDateString("tr-TR")
        : item.created_at
        ? new Date(item.created_at).toLocaleDateString("tr-TR")
        : "Tarihsiz";

      grouped[date] = (grouped[date] || 0) + 1;
    });

    const entries = Object.entries(grouped).reverse().slice(-7);
    return {
      categories: entries.map(([date]) => date),
      values: entries.map(([, value]) => value),
    };
  }, [filteredLogs]);

  const stockSummary = useMemo(() => {
    const grouped = {};

    filteredSteel.forEach((item) => {
      const name = item.name || "Adsız";
      const qty = Number(item.quantity || 0);

      if (!grouped[name]) grouped[name] = 0;
      if (item.type === "çıkış") grouped[name] -= qty;
      else grouped[name] += qty;
    });

    return Object.entries(grouped).map(([name, quantity]) => ({
      name,
      quantity,
      threshold: stockThresholds[name] ?? 10,
    }));
  }, [filteredSteel]);

  const criticalStocks = useMemo(
    () => stockSummary.filter((item) => item.quantity <= item.threshold),
    [stockSummary]
  );

  const galleryImages = useMemo(() => {
    return filteredLogs.filter((item) => item.image_url).slice(0, 12);
  }, [filteredLogs]);

  function exportExcel() {
    const data = [
      ...filteredProductions.map((item) => ({
        Tür: "İmalat",
        Ayak: item.pier || "P1",
        Ad: item.name,
        Miktar: Number(item.quantity || 0),
        Ek: "",
        Tarih: item.created_at || "",
      })),
      ...filteredSteel.map((item) => ({
        Tür: "Demir Stok",
        Ayak: item.pier || "P1",
        Ad: item.name,
        Miktar: Number(item.quantity || 0),
        Ek: item.type || "",
        Tarih: item.created_at || "",
      })),
      ...filteredProgress.map((item) => ({
        Tür: "Hakediş",
        Ayak: item.pier || "P1",
        Ad: item.item,
        Miktar: Number(item.completed_quantity || 0),
        Ek: `%${
          Number(item.total_quantity || 0)
            ? (
                (Number(item.completed_quantity || 0) /
                  Number(item.total_quantity || 0)) *
                100
              ).toFixed(1)
            : "0"
        }`,
        Tarih: item.created_at || "",
      })),
      ...filteredLogs.map((item) => ({
        Tür: "Günlük Rapor",
        Ayak: item.pier || "P1",
        Ad: item.description || "",
        Miktar: "",
        Ek: `Tarih: ${item.report_date || ""} | Hava: ${item.weather || ""} | Ekip: ${item.team || ""} | Engel: ${item.issue || ""} | Foto: ${item.image_url || ""}`,
        Tarih: item.created_at || "",
      })),
      ...stockSummary.map((item) => ({
        Tür: "Stok Özeti",
        Ayak: selectedPier,
        Ad: item.name,
        Miktar: item.quantity,
        Ek: `Kritik Seviye: ${item.threshold}`,
        Tarih: "",
      })),
    ];

    const worksheet = XLSX.utils.json_to_sheet(data);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, selectedPier);
    XLSX.writeFile(workbook, `santiye_${selectedPier}.xlsx`);
  }

  function exportDailyReportPDF() {
    const doc = new jsPDF();

    doc.setFontSize(18);
    doc.text(`TAŞKÖPRÜ VİYADÜK ŞANTİYESİ`, 14, 15);
    doc.setFontSize(13);
    doc.text(`Günlük Şantiye Raporu - ${selectedPier}`, 14, 23);
    doc.setFontSize(10);
    doc.text(
      `Oluşturma Tarihi: ${new Date().toLocaleDateString("tr-TR")}`,
      14,
      30
    );

    autoTable(doc, {
      startY: 36,
      head: [["Tarih", "Hava", "Ekip", "Yapılan İş", "Engel / Aksama", "Foto"]],
      body: filteredLogs.map((item) => [
        item.report_date || "",
        item.weather || "",
        item.team || "",
        item.description || "",
        item.issue || "",
        item.image_url ? "Var" : "Yok",
      ]),
      styles: {
        fontSize: 8,
        cellPadding: 2,
        overflow: "linebreak",
      },
      headStyles: {
        fillColor: [17, 24, 39],
      },
      alternateRowStyles: {
        fillColor: [248, 250, 252],
      },
      columnStyles: {
        0: { cellWidth: 20 },
        1: { cellWidth: 20 },
        2: { cellWidth: 25 },
        3: { cellWidth: 45 },
        4: { cellWidth: 45 },
        5: { cellWidth: 15 },
      },
    });

    doc.save(`gunluk_rapor_${selectedPier}.pdf`);
  }

  const styles = {
    page: {
      padding: 20,
      maxWidth: 1360,
      margin: "0 auto",
      fontFamily: "Arial, sans-serif",
      background: "#f3f4f6",
      minHeight: "100vh",
      color: "#111827",
    },
    titleWrap: {
      display: "flex",
      justifyContent: "space-between",
      alignItems: "center",
      flexWrap: "wrap",
      gap: 12,
      marginBottom: 20,
    },
    title: {
      fontSize: 30,
      fontWeight: 700,
      margin: 0,
    },
    subTitle: {
      fontSize: 14,
      color: "#6b7280",
      marginTop: 6,
    },
    topBar: {
      display: "flex",
      gap: 10,
      flexWrap: "wrap",
      marginBottom: 20,
    },
    button: {
      padding: "10px 14px",
      border: "none",
      borderRadius: 10,
      cursor: "pointer",
      background: "#111827",
      color: "white",
      fontWeight: 600,
      boxShadow: "0 2px 6px rgba(0,0,0,0.12)",
    },
    secondaryButton: {
      padding: "8px 12px",
      border: "none",
      borderRadius: 8,
      cursor: "pointer",
      background: "#2563eb",
      color: "white",
      fontWeight: 600,
      marginRight: 8,
    },
    dangerButton: {
      padding: "8px 12px",
      border: "none",
      borderRadius: 8,
      cursor: "pointer",
      background: "#dc2626",
      color: "white",
      fontWeight: 600,
    },
    grayButton: {
      padding: "10px 14px",
      border: "none",
      borderRadius: 10,
      cursor: "pointer",
      background: "#6b7280",
      color: "white",
      fontWeight: 600,
      marginLeft: 8,
    },
    select: {
      padding: "10px 12px",
      borderRadius: 10,
      border: "1px solid #d1d5db",
      minWidth: 100,
      background: "white",
    },
    card: {
      background: "white",
      borderRadius: 16,
      padding: 18,
      marginBottom: 16,
      boxShadow: "0 4px 14px rgba(0,0,0,0.08)",
      border: "1px solid #e5e7eb",
    },
    warningCard: {
      background: "#fff7ed",
      border: "1px solid #fdba74",
      borderRadius: 16,
      padding: 18,
      marginBottom: 16,
      boxShadow: "0 4px 14px rgba(0,0,0,0.08)",
    },
    infoCard: {
      background: "#ecfeff",
      border: "1px solid #67e8f9",
      borderRadius: 16,
      padding: 18,
      marginBottom: 16,
      boxShadow: "0 4px 14px rgba(0,0,0,0.08)",
    },
    editCard: {
      background: "#eff6ff",
      border: "1px solid #93c5fd",
      borderRadius: 16,
      padding: 18,
      marginBottom: 16,
      boxShadow: "0 4px 14px rgba(0,0,0,0.08)",
    },
    dangerText: {
      color: "#b45309",
      fontWeight: 700,
    },
    statsGrid: {
      display: "grid",
      gridTemplateColumns: "repeat(auto-fit, minmax(210px, 1fr))",
      gap: 14,
      marginBottom: 18,
    },
    statBox: {
      background: "linear-gradient(135deg, #eef2ff, #f8fafc)",
      borderRadius: 14,
      padding: 16,
      border: "1px solid #e5e7eb",
    },
    formGrid: {
      display: "grid",
      gridTemplateColumns: "repeat(auto-fit, minmax(220px, 1fr))",
      gap: 10,
      alignItems: "center",
    },
    input: {
      width: "100%",
      padding: "10px 12px",
      borderRadius: 10,
      border: "1px solid #d1d5db",
      boxSizing: "border-box",
      background: "white",
    },
    textarea: {
      width: "100%",
      padding: "10px 12px",
      borderRadius: 10,
      border: "1px solid #d1d5db",
      boxSizing: "border-box",
      minHeight: 110,
      background: "white",
    },
    listItem: {
      padding: "10px 12px",
      borderBottom: "1px solid #f1f5f9",
    },
    subtitle: {
      marginTop: 0,
      marginBottom: 14,
      fontSize: 20,
      fontWeight: 700,
    },
    dashboardGrid: {
      display: "grid",
      gridTemplateColumns: "2fr 1fr",
      gap: 16,
      alignItems: "start",
    },
    loginWrap: {
      minHeight: "100vh",
      display: "flex",
      justifyContent: "center",
      alignItems: "center",
      background: "#f3f4f6",
      padding: 20,
    },
    loginCard: {
      width: "100%",
      maxWidth: 420,
      background: "white",
      borderRadius: 16,
      padding: 24,
      boxShadow: "0 8px 24px rgba(0,0,0,0.12)",
    },
    image: {
      width: "100%",
      maxWidth: 260,
      borderRadius: 10,
      marginTop: 10,
      border: "1px solid #ddd",
      objectFit: "cover",
      cursor: "pointer",
    },
    galleryGrid: {
      display: "grid",
      gridTemplateColumns: "repeat(auto-fit, minmax(180px, 1fr))",
      gap: 12,
    },
    galleryItem: {
      background: "#fff",
      border: "1px solid #e5e7eb",
      borderRadius: 14,
      overflow: "hidden",
      boxShadow: "0 2px 8px rgba(0,0,0,0.06)",
    },
    galleryImage: {
      width: "100%",
      height: 150,
      objectFit: "cover",
      display: "block",
      cursor: "pointer",
    },
    galleryMeta: {
      padding: 10,
      fontSize: 13,
      lineHeight: 1.4,
    },
    pill: {
      display: "inline-block",
      padding: "4px 8px",
      borderRadius: 999,
      background: "#e0f2fe",
      color: "#0369a1",
      fontSize: 12,
      fontWeight: 700,
      marginBottom: 8,
    },
    actionRow: {
      marginTop: 10,
      display: "flex",
      gap: 8,
      flexWrap: "wrap",
    },
    tableWrap: {
      overflowX: "auto",
      border: "1px solid #e5e7eb",
      borderRadius: 14,
    },
    table: {
      width: "100%",
      borderCollapse: "collapse",
      minWidth: 900,
      background: "white",
    },
    th: {
      textAlign: "left",
      padding: "12px 10px",
      background: "#f8fafc",
      borderBottom: "1px solid #e5e7eb",
      fontSize: 13,
      fontWeight: 700,
      color: "#374151",
      whiteSpace: "nowrap",
    },
    td: {
      padding: "12px 10px",
      borderBottom: "1px solid #f1f5f9",
      fontSize: 14,
      verticalAlign: "top",
    },
    smallImage: {
      width: 70,
      height: 50,
      objectFit: "cover",
      borderRadius: 8,
      border: "1px solid #ddd",
      cursor: "pointer",
    },
    refreshText: {
      fontSize: 13,
      color: "#6b7280",
      marginTop: 6,
    },
    modalBackdrop: {
      position: "fixed",
      inset: 0,
      background: "rgba(0,0,0,0.75)",
      display: "flex",
      alignItems: "center",
      justifyContent: "center",
      zIndex: 9999,
      padding: 20,
    },
    modalImage: {
      maxWidth: "95vw",
      maxHeight: "90vh",
      borderRadius: 12,
      boxShadow: "0 10px 30px rgba(0,0,0,0.4)",
      background: "white",
    },
    closeButton: {
      position: "absolute",
      top: 20,
      right: 20,
      border: "none",
      background: "#111827",
      color: "white",
      borderRadius: 999,
      width: 40,
      height: 40,
      cursor: "pointer",
      fontSize: 18,
      fontWeight: 700,
    },
  };

  if (!session) {
    return (
      <div style={styles.loginWrap}>
        <div style={styles.loginCard}>
          <div style={{ fontSize: 26, fontWeight: 700, marginBottom: 16 }}>
            Şantiye Giriş
          </div>

          <input
            style={{ ...styles.input, marginBottom: 12 }}
            placeholder="E-posta"
            value={email}
            onChange={(e) => setEmail(e.target.value)}
          />

          <input
            style={{ ...styles.input, marginBottom: 12 }}
            placeholder="Şifre"
            type="password"
            value={password}
            onChange={(e) => setPassword(e.target.value)}
          />

          <button style={{ ...styles.button, width: "100%" }} onClick={signIn}>
            Giriş Yap
          </button>

          {loginError ? (
            <div style={{ color: "red", marginTop: 12 }}>{loginError}</div>
          ) : null}
        </div>
      </div>
    );
  }

  return (
    <div style={styles.page}>
      {selectedImage ? (
        <div style={styles.modalBackdrop} onClick={() => setSelectedImage("")}>
          <button style={styles.closeButton} onClick={() => setSelectedImage("")}>
            ×
          </button>
          <img
            src={selectedImage}
            alt="Büyük görsel"
            style={styles.modalImage}
            onClick={(e) => e.stopPropagation()}
          />
        </div>
      ) : null}

      <div style={styles.titleWrap}>
        <div>
          <h1 style={styles.title}>🏗️ Şantiye Takip Sistemi</h1>
          <div style={styles.subTitle}>
            Ayak bazlı üretim, stok, günlük rapor ve saha görselleri
          </div>
          <div style={styles.refreshText}>
            Son yenileme: <strong>{lastRefresh || "-"}</strong>
          </div>
        </div>
      </div>

      <div style={styles.topBar}>
        <div style={{ ...styles.card, padding: 10, marginBottom: 0 }}>
          Giriş: <strong>{session.user.email}</strong> | Yetki:{" "}
          <strong>{userRole}</strong>
        </div>

        <select
          style={styles.select}
          value={selectedPier}
          onChange={(e) => setSelectedPier(e.target.value)}
        >
          <option value="P1">P1</option>
          <option value="P2">P2</option>
          <option value="P3">P3</option>
          <option value="P4">P4</option>
        </select>

        <button style={styles.button} onClick={loadData}>
          Verileri Yenile
        </button>

        <button style={styles.button} onClick={exportExcel}>
          Excel İndir
        </button>

        <button style={styles.button} onClick={exportDailyReportPDF}>
          Günlük Rapor PDF
        </button>

        <button style={styles.button} onClick={signOut}>
          Çıkış Yap
        </button>
      </div>

      <div style={styles.infoCard}>
        <h2 style={{ marginTop: 0 }}>
          📌 Bugün Yapılan İşler Özeti - {selectedPier}
        </h2>

        <div style={styles.listItem}>
          Bugünkü İmalat Toplamı: {todayProductionTotal}
        </div>

        <div style={styles.listItem}>
          Bugünkü Rapor Sayısı: {todayLogs.length}
        </div>

        <div style={styles.listItem}>
          Bu Hafta Rapor Sayısı: {thisWeekLogsCount}
        </div>

        <div style={styles.listItem}>
          Son Durum:
          {todayWorkSummary.length > 0 ? (
            <ul style={{ marginTop: 8, paddingLeft: 20 }}>
              {todayWorkSummary.map((item, index) => (
                <li key={index}>{item}</li>
              ))}
            </ul>
          ) : (
            <span> Bugün için kayıt yok.</span>
          )}
        </div>
      </div>

      {editingLogId && canEdit ? (
        <div style={styles.editCard}>
          <strong>Şu anda düzenleme modundasın.</strong> Kaydedersen mevcut rapor güncellenir.
          <button style={styles.grayButton} onClick={resetLogForm}>
            Düzenlemeyi İptal Et
          </button>
        </div>
      ) : null}

      {criticalStocks.length > 0 && (
        <div style={styles.warningCard}>
          <h2 style={{ marginTop: 0 }}>⚠ Kritik Stok Uyarısı - {selectedPier}</h2>
          {criticalStocks.map((item) => (
            <div key={item.name} style={styles.listItem}>
              <span style={styles.dangerText}>{item.name}</span> — Mevcut:{" "}
              {item.quantity} / Kritik Seviye: {item.threshold}
            </div>
          ))}
        </div>
      )}

      <div style={styles.statsGrid}>
        <div style={styles.statBox}>
          <strong>Toplam İmalat</strong>
          <div style={{ fontSize: 26, marginTop: 8 }}>{totalProduction}</div>
        </div>

        <div style={styles.statBox}>
          <strong>Toplam Demir Stok</strong>
          <div style={{ fontSize: 26, marginTop: 8 }}>{totalSteel}</div>
        </div>

        <div style={styles.statBox}>
          <strong>Ortalama Hakediş</strong>
          <div style={{ fontSize: 26, marginTop: 8 }}>%{avgProgress}</div>
        </div>

        <div style={styles.statBox}>
          <strong>Seçili Ayakta Günlük Rapor</strong>
          <div style={{ fontSize: 26, marginTop: 8 }}>{filteredLogs.length}</div>
        </div>

        <div style={styles.statBox}>
          <strong>En Çok İmalat Yapılan Ayak</strong>
          <div style={{ fontSize: 26, marginTop: 8 }}>{maxPierData.pier}</div>
          <div style={{ marginTop: 4 }}>{maxPierData.total}</div>
        </div>

        <div style={styles.statBox}>
          <strong>Fotoğraflı Rapor</strong>
          <div style={{ fontSize: 26, marginTop: 8 }}>{galleryImages.length}</div>
        </div>
      </div>

      <div style={styles.card}>
        <h2 style={styles.subtitle}>🔎 Rapor Filtreleme</h2>
        <div style={styles.formGrid}>
          <input
            type="date"
            style={styles.input}
            value={filterStartDate}
            onChange={(e) => setFilterStartDate(e.target.value)}
          />

          <input
            type="date"
            style={styles.input}
            value={filterEndDate}
            onChange={(e) => setFilterEndDate(e.target.value)}
          />

          <input
            style={styles.input}
            placeholder="Metin ara..."
            value={filterText}
            onChange={(e) => setFilterText(e.target.value)}
          />

          <label style={{ display: "flex", alignItems: "center", gap: 8 }}>
            <input
              type="checkbox"
              checked={onlyWithImage}
              onChange={(e) => setOnlyWithImage(e.target.checked)}
            />
            Sadece fotoğraflı
          </label>
        </div>
      </div>

      <div style={styles.dashboardGrid}>
        <div>
          <div style={styles.card}>
            <h2 style={styles.subtitle}>Tüm Ayaklar Karşılaştırma</h2>
            <Chart
              options={{
                chart: { id: "tum-ayaklar", toolbar: { show: false } },
                xaxis: { categories: ["P1", "P2", "P3", "P4"] },
                dataLabels: { enabled: true },
              }}
              series={[{ name: "Toplam İmalat", data: pierTotals }]}
              type="bar"
              width="100%"
              height={320}
            />
          </div>

          <div style={styles.card}>
            <h2 style={styles.subtitle}>{selectedPier} Detay Grafiği</h2>
            <Chart
              options={{
                chart: { id: "imalat-grafigi", toolbar: { show: false } },
                xaxis: { categories: filteredProductions.map((p) => p.name) },
                dataLabels: { enabled: true },
              }}
              series={[
                {
                  name: `${selectedPier} İmalat`,
                  data: filteredProductions.map((p) => Number(p.quantity || 0)),
                },
              ]}
              type="bar"
              width="100%"
              height={320}
            />
          </div>

          <div style={styles.card}>
            <h2 style={styles.subtitle}>{selectedPier} Günlük İmalat</h2>
            <Chart
              options={{
                chart: { id: "gunluk-imalat", toolbar: { show: false } },
                xaxis: { categories: productionByDay.categories },
                stroke: { curve: "smooth" },
                dataLabels: { enabled: true },
              }}
              series={[{ name: "Günlük İmalat", data: productionByDay.values }]}
              type="line"
              width="100%"
              height={320}
            />
          </div>

          <div style={styles.card}>
            <h2 style={styles.subtitle}>{selectedPier} Günlük Rapor Sayısı</h2>
            <Chart
              options={{
                chart: { id: "gunluk-rapor", toolbar: { show: false } },
                xaxis: { categories: logsByDay.categories },
                stroke: { curve: "smooth" },
                dataLabels: { enabled: true },
              }}
              series={[{ name: "Rapor Adedi", data: logsByDay.values }]}
              type="line"
              width="100%"
              height={320}
            />
          </div>

          <div style={styles.card}>
            <h2 style={styles.subtitle}>📷 Günlük Rapor Galerisi - {selectedPier}</h2>
            {galleryImages.length === 0 ? (
              <p>Görsel yok</p>
            ) : (
              <div style={styles.galleryGrid}>
                {galleryImages.map((item) => (
                  <div key={item.id} style={styles.galleryItem}>
                    <img
                      src={item.image_url}
                      alt="Saha görseli"
                      style={styles.galleryImage}
                      onClick={() => setSelectedImage(item.image_url)}
                      onError={(e) => {
                        e.currentTarget.style.display = "none";
                      }}
                    />
                    <div style={styles.galleryMeta}>
                      <div style={styles.pill}>{item.report_date || "-"}</div>
                      <div><strong>Hava:</strong> {item.weather || "-"}</div>
                      <div><strong>Ekip:</strong> {item.team || "-"}</div>
                      <div style={{ marginTop: 6 }}>
                        <strong>Açıklama:</strong> {item.description || "-"}
                      </div>
                    </div>
                  </div>
                ))}
              </div>
            )}
          </div>

          <div style={styles.card}>
            <h2 style={styles.subtitle}>📋 Günlük Rapor Tablosu - {selectedPier}</h2>
            <div style={styles.tableWrap}>
              <table style={styles.table}>
                <thead>
                  <tr>
                    <th style={styles.th}>Tarih</th>
                    <th style={styles.th}>Hava</th>
                    <th style={styles.th}>Ekip</th>
                    <th style={styles.th}>Yapılan İş</th>
                    <th style={styles.th}>Engel / Aksama</th>
                    <th style={styles.th}>Foto</th>
                    {canEdit ? <th style={styles.th}>İşlem</th> : null}
                  </tr>
                </thead>
                <tbody>
                  {filteredLogs.length === 0 ? (
                    <tr>
                      <td style={styles.td} colSpan={canEdit ? 7 : 6}>
                        Kayıt yok
                      </td>
                    </tr>
                  ) : (
                    filteredLogs.map((item) => (
                      <tr key={item.id}>
                        <td style={styles.td}>{item.report_date || "-"}</td>
                        <td style={styles.td}>{item.weather || "-"}</td>
                        <td style={styles.td}>{item.team || "-"}</td>
                        <td style={styles.td}>{item.description || "-"}</td>
                        <td style={styles.td}>{item.issue || "-"}</td>
                        <td style={styles.td}>
                          {item.image_url ? (
                            <img
                              src={item.image_url}
                              alt="Rapor foto"
                              style={styles.smallImage}
                              onClick={() => setSelectedImage(item.image_url)}
                              onError={(e) => {
                                e.currentTarget.style.display = "none";
                              }}
                            />
                          ) : (
                            "Yok"
                          )}
                        </td>
                        {canEdit ? (
                          <td style={styles.td}>
                            <button
                              style={styles.secondaryButton}
                              onClick={() => startEditLog(item)}
                            >
                              Düzenle
                            </button>
                            <button
                              style={styles.dangerButton}
                              onClick={() => deleteLog(item.id)}
                            >
                              Sil
                            </button>
                          </td>
                        ) : null}
                      </tr>
                    ))
                  )}
                </tbody>
              </table>
            </div>
          </div>
        </div>

        <div>
          <div style={styles.card}>
            <h2 style={styles.subtitle}>Son 5 Günlük Rapor - {selectedPier}</h2>
            {recentLogs.length === 0 ? (
              <p>Rapor yok</p>
            ) : (
              recentLogs.map((item) => (
                <div key={item.id} style={styles.listItem}>
                  <strong>Tarih:</strong> {item.report_date || "-"}
                  <div><strong>Hava:</strong> {item.weather || "-"}</div>
                  <div><strong>Ekip:</strong> {item.team || "-"}</div>
                  <div style={{ marginTop: 6 }}>
                    <strong>Yapılan İş:</strong> {item.description || "-"}
                  </div>
                  <div style={{ marginTop: 6 }}>
                    <strong>Engel/Aksama:</strong> {item.issue || "-"}
                  </div>
                  {item.image_url ? (
                    <img
                      src={item.image_url}
                      alt="Rapor görseli"
                      style={styles.image}
                      onClick={() => setSelectedImage(item.image_url)}
                      onError={(e) => {
                        e.currentTarget.style.display = "none";
                      }}
                    />
                  ) : (
                    <div style={{ marginTop: 8, color: "#999" }}>Fotoğraf yok</div>
                  )}

                  {canEdit && (
                    <div style={styles.actionRow}>
                      <button
                        style={styles.secondaryButton}
                        onClick={() => startEditLog(item)}
                      >
                        Düzenle
                      </button>
                      <button
                        style={styles.dangerButton}
                        onClick={() => deleteLog(item.id)}
                      >
                        Sil
                      </button>
                    </div>
                  )}
                </div>
              ))
            )}
          </div>

          <div style={styles.card}>
            <h2 style={styles.subtitle}>Stok Özeti - {selectedPier}</h2>
            {stockSummary.length === 0 ? (
              <p>Stok kaydı yok</p>
            ) : (
              stockSummary.map((item) => (
                <div key={item.name} style={styles.listItem}>
                  <strong>{item.name}</strong> — Mevcut: {item.quantity}
                  {item.quantity <= item.threshold && (
                    <span style={{ color: "#b45309", fontWeight: 700 }}>
                      {" "}
                      (Kritik)
                    </span>
                  )}
                </div>
              ))
            )}
          </div>

          <div style={styles.card}>
            <h2 style={styles.subtitle}>Hızlı Özet - {selectedPier}</h2>
            <div style={styles.listItem}>
              İmalat Kaydı: {filteredProductions.length}
            </div>
            <div style={styles.listItem}>
              Demir Hareketi: {filteredSteel.length}
            </div>
            <div style={styles.listItem}>
              Hakediş Kalemi: {filteredProgress.length}
            </div>
            <div style={styles.listItem}>
              Günlük Rapor: {filteredLogs.length}
            </div>
            <div style={styles.listItem}>
              Fotoğraflı Rapor: {galleryImages.length}
            </div>
          </div>
        </div>
      </div>

      {canEdit ? (
        <>
          <div style={styles.card}>
            <h2 style={styles.subtitle}>İmalat Girişi - {selectedPier}</h2>
            <div style={styles.formGrid}>
              <input
                style={styles.input}
                placeholder="İmalat adı"
                value={productionName}
                onChange={(e) => setProductionName(e.target.value)}
              />
              <input
                style={styles.input}
                placeholder="Miktar"
                type="number"
                value={productionQty}
                onChange={(e) => setProductionQty(e.target.value)}
              />
              <button style={styles.button} onClick={addProduction}>
                Kaydet
              </button>
            </div>
          </div>

          <div style={styles.card}>
            <h2 style={styles.subtitle}>Demir Stok Girişi - {selectedPier}</h2>
            <div style={styles.formGrid}>
              <input
                style={styles.input}
                placeholder="Malzeme adı"
                value={steelName}
                onChange={(e) => setSteelName(e.target.value)}
              />
              <input
                style={styles.input}
                placeholder="Miktar"
                type="number"
                value={steelQty}
                onChange={(e) => setSteelQty(e.target.value)}
              />
              <select
                style={styles.select}
                value={steelType}
                onChange={(e) => setSteelType(e.target.value)}
              >
                <option value="giriş">giriş</option>
                <option value="çıkış">çıkış</option>
              </select>
              <button style={styles.button} onClick={addSteel}>
                Kaydet
              </button>
            </div>
          </div>

          <div style={styles.card}>
            <h2 style={styles.subtitle}>Hakediş Girişi - {selectedPier}</h2>
            <div style={styles.formGrid}>
              <input
                style={styles.input}
                placeholder="İş kalemi"
                value={progressItem}
                onChange={(e) => setProgressItem(e.target.value)}
              />
              <input
                style={styles.input}
                placeholder="Toplam miktar"
                type="number"
                value={progressTotal}
                onChange={(e) => setProgressTotal(e.target.value)}
              />
              <input
                style={styles.input}
                placeholder="Tamamlanan miktar"
                type="number"
                value={progressDone}
                onChange={(e) => setProgressDone(e.target.value)}
              />
              <button style={styles.button} onClick={addProgress}>
                Kaydet
              </button>
            </div>
          </div>

          <div style={styles.card}>
            <h2 style={styles.subtitle}>
              {editingLogId
                ? `Günlük Rapor Düzenle - ${selectedPier}`
                : `Gelişmiş Günlük Rapor - ${selectedPier}`}
            </h2>

            <div style={styles.formGrid}>
              <input
                style={styles.input}
                type="date"
                value={reportDate}
                onChange={(e) => setReportDate(e.target.value)}
              />
              <input
                style={styles.input}
                placeholder="Hava durumu"
                value={weather}
                onChange={(e) => setWeather(e.target.value)}
              />
              <input
                style={styles.input}
                placeholder="Çalışan ekip"
                value={team}
                onChange={(e) => setTeam(e.target.value)}
              />
            </div>

            <div style={{ marginTop: 12 }}>
              <textarea
                style={styles.textarea}
                value={logText}
                onChange={(e) => setLogText(e.target.value)}
                placeholder="Bugün yapılan işleri yaz"
              />
            </div>

            <div style={{ marginTop: 12 }}>
              <textarea
                style={styles.textarea}
                value={issue}
                onChange={(e) => setIssue(e.target.value)}
                placeholder="Engel / aksama / iş durduran sebepler"
              />
            </div>

            {existingImageUrl ? (
              <div style={{ marginTop: 12 }}>
                <div style={{ marginBottom: 6, fontWeight: 700 }}>
                  Mevcut fotoğraf
                </div>
                <img
                  src={existingImageUrl}
                  alt="Mevcut rapor görseli"
                  style={styles.image}
                  onClick={() => setSelectedImage(existingImageUrl)}
                />
              </div>
            ) : null}

            <div style={{ marginTop: 12 }}>
              <input
                style={styles.input}
                type="file"
                accept="image/*"
                onChange={(e) => setReportImage(e.target.files?.[0] || null)}
              />
              {reportImage ? (
                <div style={{ marginTop: 8, fontSize: 14, color: "#555" }}>
                  Seçilen dosya: {reportImage.name}
                </div>
              ) : null}
            </div>

            <div style={{ marginTop: 10 }}>
              <button style={styles.button} onClick={addLog} disabled={isSavingLog}>
                {isSavingLog
                  ? "Kaydediliyor..."
                  : editingLogId
                  ? "Güncellemeyi Kaydet"
                  : "Raporu Kaydet"}
              </button>

              {editingLogId ? (
                <button style={styles.grayButton} onClick={resetLogForm}>
                  İptal
                </button>
              ) : null}
            </div>
          </div>
        </>
      ) : null}
    </div>
  );
}
