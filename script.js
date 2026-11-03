/* ═══════════════════════════════════════════════════════════════
   MVIKAS DASHBOARD — October 2026 Executive Analytics
   Updated from MVLOAD October.xlsx
   All weights strictly in Tonnes (T) · Real Finance Data from Sheets
   ═══════════════════════════════════════════════════════════════ */

// Palette — Dark Base + MVIKAS Orange Accent
const C = {
  orange: '#f58220',
  orangeDim: '#d96d10',
  blue: '#3b82f6',
  green: '#10b981',
  yellow: '#eab308',
  bg: '#0a0a0a',
  s1: '#141414',
  s2: '#1a1a1a',
  s3: '#202020',
  s4: '#262626',
  border: '#2a2a2a',
  divider: '#3a3a3a',
  text: '#f5f5f5',
  text2: '#b8b8b8',
  text3: '#8b8b8b',
  text4: '#5a5a5a',
  grid: '#252525'
};
const orangeA = (a) => `rgba(245,130,32,${a})`;

// DOM Helper
const $ = id => document.getElementById(id);

// Safe Chart.js wrapper
function safeChart(ctx, config) {
  if (typeof Chart === 'undefined') { console.warn('Chart.js not loaded — skipping chart render'); return null; }
  if (!ctx) return null;
  try { return new Chart(ctx, config); }
  catch (e) { console.error('Chart render failed:', e); return null; }
}

// Chart defaults (Dark theme)
if (typeof Chart !== 'undefined') {
  const CD = Chart.defaults;
  CD.color = C.text3;
  CD.font.family = 'Inter, -apple-system, system-ui, sans-serif';
  CD.font.size = 11;
  CD.plugins.tooltip.backgroundColor = C.s3;
  CD.plugins.tooltip.borderColor = C.border;
  CD.plugins.tooltip.borderWidth = 1;
  CD.plugins.tooltip.padding = 12;
  CD.plugins.tooltip.cornerRadius = 8;
  CD.plugins.tooltip.titleColor = C.orange;
  CD.plugins.tooltip.titleFont = { size: 11, weight: '700' };
  CD.plugins.tooltip.bodyColor = C.text;
  CD.plugins.tooltip.bodyFont = { size: 12, weight: '500' };
  CD.plugins.tooltip.displayColors = false;
  CD.scale.grid.color = C.grid;
  CD.scale.grid.drawBorder = false;
  CD.scale.ticks.color = C.text3;
}

// Counter Animation
function animateCounters() {
  document.querySelectorAll('.counter').forEach(el => {
    const t = +el.dataset.target, dec = +(el.dataset.decimals || 0), dur = 1800, st = performance.now();
    if (isNaN(t)) return;
    (function tick(now) {
      const p = Math.min((now - st) / dur, 1);
      const v = t * (1 - Math.pow(1 - p, 3));
      el.textContent = dec > 0 ? v.toFixed(dec) : Math.round(v).toLocaleString('en-IN');
      if (p < 1) requestAnimationFrame(tick);
    })(st);
  });
}

// Tab Switching
function switchTab(id, btn) {
  document.querySelectorAll('.tab').forEach(t => {
    t.classList.remove('active');
    t.setAttribute('aria-selected', 'false');
  });
  document.querySelectorAll('.panel').forEach(p => p.classList.remove('active'));
  const targetPanel = document.getElementById('tab-' + id);
  if (targetPanel) targetPanel.classList.add('active');
  if (btn) {
    btn.classList.add('active');
    btn.setAttribute('aria-selected', 'true');
  }
}
window.switchTab = switchTab;

// ═══════════════════════════════════════════════════════════════
// DATASETS LOADED FROM MVLOAD OCTOBER.XLSX
// ═══════════════════════════════════════════════════════════════

// 1. Client Monthly Performance (Total_Tonnage_this_month sheet)
// All targets and achieved weights are strictly stored in Tonnes (T)
const clients = [
  {
    "name": "Loom Solar",
    "target": 13.61,
    "achieved": 13.0,
    "taxable": 214426.92,
    "finTarget": 300000.0,
    "activeDays": 20,
    "person": "Deepak Sharma",
    "periods": {
      "first10": {
        "target": 48.61,
        "achieved": 5.0
      },
      "mid10": {
        "target": 34.03,
        "achieved": 2.75
      },
      "last10": {
        "target": 42.54,
        "achieved": 5.24
      }
    }
  },
  {
    "name": "Bombax",
    "target": 71.28,
    "achieved": 67.9,
    "taxable": 866797.12,
    "finTarget": 1000000.0,
    "activeDays": 20,
    "person": "Sangeet Dhasmana",
    "periods": {
      "first10": {
        "target": 222.74,
        "achieved": 39.65
      },
      "mid10": {
        "target": 215.99,
        "achieved": 14.3
      },
      "last10": {
        "target": 203.65,
        "achieved": 13.95
      }
    }
  },
  {
    "name": "Sunrise Sports India Pvt Ltd",
    "target": 71.43,
    "achieved": 94.87,
    "taxable": 2352609.11,
    "finTarget": 2000000.0,
    "activeDays": 20,
    "person": "Deepak Sharma",
    "periods": {
      "first10": {
        "target": 23.81,
        "achieved": 36.69
      },
      "mid10": {
        "target": 23.81,
        "achieved": 9.23
      },
      "last10": {
        "target": 23.81,
        "achieved": 48.96
      }
    }
  },
  {
    "name": "Haier CCR",
    "target": 27.97,
    "achieved": 30.25,
    "taxable": 691619.92,
    "finTarget": 400000.0,
    "activeDays": 20,
    "person": "Deepak Sharma",
    "periods": {
      "first10": {
        "target": 69.93,
        "achieved": 6.51
      },
      "mid10": {
        "target": 107.58,
        "achieved": 12.22
      },
      "last10": {
        "target": 84.76,
        "achieved": 11.52
      }
    }
  },
  {
    "name": "Kumar Services",
    "target": 14.46,
    "achieved": 14.55,
    "taxable": 130561.24,
    "finTarget": 150000.0,
    "activeDays": 20,
    "person": "Deepak Sharma",
    "periods": {
      "first10": {
        "target": 49.88,
        "achieved": 4.77
      },
      "mid10": {
        "target": 43.83,
        "achieved": 2.8
      },
      "last10": {
        "target": 38.07,
        "achieved": 6.31
      }
    }
  },
  {
    "name": "Sukuga",
    "target": 14.45,
    "achieved": 16.8,
    "taxable": 291127.15,
    "finTarget": 200000.0,
    "activeDays": 20,
    "person": "Deepak Sharma",
    "periods": {
      "first10": {
        "target": 68.81,
        "achieved": 5.89
      },
      "mid10": {
        "target": 49.83,
        "achieved": 2.53
      },
      "last10": {
        "target": 29.49,
        "achieved": 8.38
      }
    }
  },
  {
    "name": "Carrier Refrigeration",
    "target": 346.53,
    "achieved": 551.36,
    "taxable": 7508689.07,
    "finTarget": 4200000.0,
    "activeDays": 20,
    "person": "Sangeet Dhasmana",
    "periods": {
      "first10": {
        "target": 1332.83,
        "achieved": 214.59
      },
      "mid10": {
        "target": 1283.46,
        "achieved": 115.84
      },
      "last10": {
        "target": 737.31,
        "achieved": 220.92
      }
    }
  },
  {
    "name": "Carrier - CTD",
    "target": 87.34,
    "achieved": 95.54,
    "taxable": 1394094.07,
    "finTarget": 1200000.0,
    "activeDays": 20,
    "person": "Deepak Sharma",
    "periods": {
      "first10": {
        "target": 281.73,
        "achieved": 38.63
      },
      "mid10": {
        "target": 249.53,
        "achieved": 28.96
      },
      "last10": {
        "target": 256.87,
        "achieved": 27.95
      }
    }
  },
  {
    "name": "Oneric Appliances",
    "target": 31.04,
    "achieved": 2.26,
    "taxable": 26859.94,
    "finTarget": 500000.0,
    "activeDays": 20,
    "person": "Deepak Sharma",
    "periods": {
      "first10": {
        "target": 72.18,
        "achieved": 0.51
      },
      "mid10": {
        "target": 114.95,
        "achieved": 0.57
      },
      "last10": {
        "target": 103.46,
        "achieved": 1.18
      }
    }
  },
  {
    "name": "Vaidrishi Laboratories",
    "target": 2.03,
    "achieved": 9.58,
    "taxable": 145726.55,
    "finTarget": 30000.0,
    "activeDays": 20,
    "person": "Sangeet Dhasmana",
    "periods": {
      "first10": {
        "target": 5.98,
        "achieved": 3.96
      },
      "mid10": {
        "target": 7.53,
        "achieved": 1.49
      },
      "last10": {
        "target": 5.21,
        "achieved": 4.13
      }
    }
  },
  {
    "name": "Abhishek Tomar",
    "target": 0.0,
    "achieved": 0.38,
    "taxable": 3664.33,
    "finTarget": 0.0,
    "activeDays": 20,
    "person": "Deepak Sharma",
    "periods": {
      "first10": {
        "target": 0.0,
        "achieved": 0.38
      },
      "mid10": {
        "target": 0.0,
        "achieved": 0.0
      },
      "last10": {
        "target": 0.0,
        "achieved": 0.0
      }
    }
  },
  {
    "name": "Cosmos Pumps Pvt. Ltd.",
    "target": 8.29,
    "achieved": 13.71,
    "taxable": 170594.41,
    "finTarget": 100000.0,
    "activeDays": 20,
    "person": "Deepak Sharma",
    "periods": {
      "first10": {
        "target": 23.69,
        "achieved": 5.6
      },
      "mid10": {
        "target": 33.17,
        "achieved": 1.88
      },
      "last10": {
        "target": 20.73,
        "achieved": 6.23
      }
    }
  },
  {
    "name": "Conficore",
    "target": 2.84,
    "achieved": 0.87,
    "taxable": 11373.24,
    "finTarget": 50000.0,
    "activeDays": 20,
    "person": "Sangeet Dhasmana",
    "periods": {
      "first10": {
        "target": 3.64,
        "achieved": 0.87
      },
      "mid10": {
        "target": 70.9,
        "achieved": 0.0
      },
      "last10": {
        "target": 15.76,
        "achieved": 0.0
      }
    }
  },
  {
    "name": "MS Medicals Pvt Ltd",
    "target": 5.77,
    "achieved": 6.36,
    "taxable": 81070.78,
    "finTarget": 75000.0,
    "activeDays": 20,
    "person": "Sangeet Dhasmana",
    "periods": {
      "first10": {
        "target": 19.23,
        "achieved": 1.26
      },
      "mid10": {
        "target": 0.3,
        "achieved": 4.42
      },
      "last10": {
        "target": 11.31,
        "achieved": 0.68
      }
    }
  },
  {
    "name": "Khatana Enterprises",
    "target": 0.0,
    "achieved": 0.42,
    "taxable": 8470.0,
    "finTarget": 0.0,
    "activeDays": 20,
    "person": "Deepak Sharma",
    "periods": {
      "first10": {
        "target": 0.0,
        "achieved": 0.42
      },
      "mid10": {
        "target": 0.0,
        "achieved": 0.0
      },
      "last10": {
        "target": 0.0,
        "achieved": 0.0
      }
    }
  },
  {
    "name": "MITRAS",
    "target": 3.74,
    "achieved": 0.89,
    "taxable": 9560.31,
    "finTarget": 50000.0,
    "activeDays": 20,
    "person": "Deepak Sharma",
    "periods": {
      "first10": {
        "target": 20.76,
        "achieved": 0.05
      },
      "mid10": {
        "target": 8.49,
        "achieved": 0.0
      },
      "last10": {
        "target": 9.83,
        "achieved": 0.84
      }
    }
  },
  {
    "name": "Droplets Engineering",
    "target": 0.0,
    "achieved": 0.43,
    "taxable": 13948.76,
    "finTarget": 0.0,
    "activeDays": 20,
    "person": "Deepak Sharma",
    "periods": {
      "first10": {
        "target": 0.0,
        "achieved": 0.0
      },
      "mid10": {
        "target": 0.0,
        "achieved": 0.0
      },
      "last10": {
        "target": 0.0,
        "achieved": 0.43
      }
    }
  },
  {
    "name": "Khusbhu Enterprises",
    "target": 0.0,
    "achieved": 0.06,
    "taxable": 727.0,
    "finTarget": 0.0,
    "activeDays": 20,
    "person": "Deepak Sharma",
    "periods": {
      "first10": {
        "target": 0.0,
        "achieved": 0.0
      },
      "mid10": {
        "target": 0.0,
        "achieved": 0.0
      },
      "last10": {
        "target": 0.0,
        "achieved": 0.06
      }
    }
  },
  {
    "name": "KTB",
    "target": 0.0,
    "achieved": 0.05,
    "taxable": 3850.0,
    "finTarget": 0.0,
    "activeDays": 20,
    "person": "Sangeet Dhasmana",
    "periods": {
      "first10": {
        "target": 0.0,
        "achieved": 0.0
      },
      "mid10": {
        "target": 0.0,
        "achieved": 0.0
      },
      "last10": {
        "target": 0.0,
        "achieved": 0.05
      }
    }
  },
  {
    "name": "Indu Sports",
    "target": 0.0,
    "achieved": 1.65,
    "taxable": 25962.53,
    "finTarget": 0.0,
    "activeDays": 20,
    "person": "Deepak Sharma",
    "periods": {
      "first10": {
        "target": 0.0,
        "achieved": 0.0
      },
      "mid10": {
        "target": 0.0,
        "achieved": 0.0
      },
      "last10": {
        "target": 0.0,
        "achieved": 1.65
      }
    }
  }
];

clients.forEach(c => {
  c.pct = c.target > 0 ? Math.round(c.achieved / c.target * 100) : (c.achieved > 0 ? 100 : 0);
  c.finPct = c.finTarget > 0 ? Math.round(c.taxable / c.finTarget * 100) : (c.taxable > 0 ? 100 : 0);
  c.avgDay = c.activeDays > 0 ? +(c.achieved / c.activeDays).toFixed(2) : 0;
  c.remaining = +Math.max(c.target - c.achieved, 0).toFixed(2);
  c.daysNeeded = c.avgDay > 0 && c.remaining > 0 ? +(c.remaining / c.avgDay).toFixed(1) : (c.remaining === 0 ? 0 : 999);
  c.status = c.pct >= 100 ? 'Target Achieved' : (c.pct >= 70 ? 'On Track' : 'Needs Focus');
});

// 2. Daily Tonnage from Sheet (Yesterday dispatches across 11 clients)
const dailyTonnage = [
  {
    "name": "Carrier Refrigeration",
    "tonnes": 33.66
  },
  {
    "name": "Sunrise Sports India Pvt Ltd",
    "tonnes": 2.87
  },
  {
    "name": "Loom Solar",
    "tonnes": 2.11
  },
  {
    "name": "Bombax",
    "tonnes": 4.19
  },
  {
    "name": "Haier CCR",
    "tonnes": 1.59
  },
  {
    "name": "Sukuga",
    "tonnes": 1.11
  },
  {
    "name": "Cosmos Pumps Pvt. Ltd.",
    "tonnes": 1.01
  },
  {
    "name": "Oneric Appliances",
    "tonnes": 0.69
  },
  {
    "name": "Vaidrishi Laboratories",
    "tonnes": 0.28
  },
  {
    "name": "Carrier - CTD",
    "tonnes": 0.64
  },
  {
    "name": "Kumar Services",
    "tonnes": 1.22
  }
];
const dailyTotalTonnage = +(dailyTonnage.reduce((a, b) => a + b.tonnes, 0)).toFixed(2);

// 3. Orders Due Tomorrow (123 Orders from Order due Tommorow sheet)
const dueOrders = [
  {
    "id": "MVS/26-27/30479",
    "customer": "Sunrise Sports India Pvt Ltd",
    "taxable": 3369.34,
    "total": 3975.82,
    "docket": "20002066286",
    "transporter": "XP INDIA",
    "edd": "2026-11-01",
    "bookingDate": "2026-10-22"
  },
  {
    "id": "MVS/26-27/30682",
    "customer": "Bombax",
    "taxable": 691.79,
    "total": 816.31,
    "docket": "1303691217",
    "transporter": "DP WORLD",
    "edd": "2026-11-01",
    "bookingDate": "2026-10-24"
  },
  {
    "id": "MVS/26-27/30721",
    "customer": "Carrier - CTD",
    "taxable": 4243.21,
    "total": 5006.99,
    "docket": "1303559618",
    "transporter": "DP WORLD",
    "edd": "2026-11-01",
    "bookingDate": "2026-10-24"
  },
  {
    "id": "MVS/26-27/30752",
    "customer": "Bombax",
    "taxable": 443.75,
    "total": 523.63,
    "docket": "2016828933",
    "transporter": "EKART",
    "edd": "2026-11-01",
    "bookingDate": "2026-10-24"
  },
  {
    "id": "MVS/26-27/30774",
    "customer": "Carrier Refrigeration",
    "taxable": 2717.41,
    "total": 3206.54,
    "docket": "20002071227",
    "transporter": "XP INDIA",
    "edd": "2026-11-01",
    "bookingDate": "2026-10-24"
  },
  {
    "id": "MVS/26-27/30785",
    "customer": "Carrier Refrigeration",
    "taxable": 1709.09,
    "total": 2016.73,
    "docket": "2016830134",
    "transporter": "EKART",
    "edd": "2026-11-01",
    "bookingDate": "2026-10-24"
  },
  {
    "id": "MVS/26-27/30791",
    "customer": "Carrier Refrigeration",
    "taxable": 2627.29,
    "total": 3100.2,
    "docket": "1303620001",
    "transporter": "DP WORLD",
    "edd": "2026-11-01",
    "bookingDate": "2026-10-24"
  },
  {
    "id": "MVS/26-27/30792",
    "customer": "Carrier Refrigeration",
    "taxable": 4755.53,
    "total": 5611.53,
    "docket": "1303620003",
    "transporter": "DP WORLD",
    "edd": "2026-11-01",
    "bookingDate": "2026-10-24"
  },
  {
    "id": "MVS/26-27/30821",
    "customer": "Carrier Refrigeration",
    "taxable": 1436.75,
    "total": 1695.37,
    "docket": "1303557984",
    "transporter": "DP WORLD",
    "edd": "2026-11-01",
    "bookingDate": "2026-10-24"
  },
  {
    "id": "MVS/26-27/30834",
    "customer": "Sunrise Sports India Pvt Ltd",
    "taxable": 611.2,
    "total": 721.22,
    "docket": "20002075192",
    "transporter": "XP INDIA",
    "edd": "2026-11-01",
    "bookingDate": "2026-10-26"
  },
  {
    "id": "MVS/26-27/30835",
    "customer": "Sunrise Sports India Pvt Ltd",
    "taxable": 613.29,
    "total": 723.68,
    "docket": "20002075193",
    "transporter": "XP INDIA",
    "edd": "2026-11-01",
    "bookingDate": "2026-10-26"
  },
  {
    "id": "MVS/26-27/30837",
    "customer": "Sunrise Sports India Pvt Ltd",
    "taxable": 4378.44,
    "total": 5166.56,
    "docket": "20002075195",
    "transporter": "XP INDIA",
    "edd": "2026-11-01",
    "bookingDate": "2026-10-26"
  },
  {
    "id": "MVS/26-27/30839",
    "customer": "Sunrise Sports India Pvt Ltd",
    "taxable": 4130.57,
    "total": 4874.07,
    "docket": "20002075309",
    "transporter": "XP INDIA",
    "edd": "2026-11-01",
    "bookingDate": "2026-10-26"
  },
  {
    "id": "MVS/26-27/30840",
    "customer": "Sunrise Sports India Pvt Ltd",
    "taxable": 4741.08,
    "total": 5594.47,
    "docket": "20002075310",
    "transporter": "XP INDIA",
    "edd": "2026-11-01",
    "bookingDate": "2026-10-26"
  },
  {
    "id": "MVS/26-27/30841",
    "customer": "Sunrise Sports India Pvt Ltd",
    "taxable": 6940.53,
    "total": 8189.83,
    "docket": "20002075311",
    "transporter": "XP INDIA",
    "edd": "2026-11-01",
    "bookingDate": "2026-10-26"
  },
  {
    "id": "MVS/26-27/30842",
    "customer": "Sunrise Sports India Pvt Ltd",
    "taxable": 630.4,
    "total": 743.87,
    "docket": "20002075324",
    "transporter": "XP INDIA",
    "edd": "2026-11-01",
    "bookingDate": "2026-10-26"
  },
  {
    "id": "MVS/26-27/30843",
    "customer": "Sunrise Sports India Pvt Ltd",
    "taxable": 742.36,
    "total": 875.98,
    "docket": "20002075312",
    "transporter": "XP INDIA",
    "edd": "2026-11-01",
    "bookingDate": "2026-10-26"
  },
  {
    "id": "MVS/26-27/30846",
    "customer": "Sunrise Sports India Pvt Ltd",
    "taxable": 3773.13,
    "total": 4452.29,
    "docket": "20002075762",
    "transporter": "XP INDIA",
    "edd": "2026-11-01",
    "bookingDate": "2026-10-26"
  },
  {
    "id": "MVS/26-27/30847",
    "customer": "Sunrise Sports India Pvt Ltd",
    "taxable": 2354.22,
    "total": 2777.98,
    "docket": "20002075763",
    "transporter": "XP INDIA",
    "edd": "2026-11-01",
    "bookingDate": "2026-10-26"
  },
  {
    "id": "MVS/26-27/30848",
    "customer": "Sunrise Sports India Pvt Ltd",
    "taxable": 6637.1,
    "total": 7831.78,
    "docket": "20002075860",
    "transporter": "XP INDIA",
    "edd": "2026-11-01",
    "bookingDate": "2026-10-26"
  },
  {
    "id": "MVS/26-27/30854",
    "customer": "Sunrise Sports India Pvt Ltd",
    "taxable": 2836.96,
    "total": 3347.61,
    "docket": "20002076309",
    "transporter": "XP INDIA",
    "edd": "2026-11-01",
    "bookingDate": "2026-10-26"
  },
  {
    "id": "MVS/26-27/30855",
    "customer": "Sunrise Sports India Pvt Ltd",
    "taxable": 1075.88,
    "total": 1269.54,
    "docket": "20002076310",
    "transporter": "XP INDIA",
    "edd": "2026-11-01",
    "bookingDate": "2026-10-26"
  },
  {
    "id": "MVS/26-27/30856",
    "customer": "Sunrise Sports India Pvt Ltd",
    "taxable": 1568.34,
    "total": 1850.64,
    "docket": "20002076393",
    "transporter": "XP INDIA",
    "edd": "2026-11-01",
    "bookingDate": "2026-10-26"
  },
  {
    "id": "MVS/26-27/30888",
    "customer": "Carrier Refrigeration",
    "taxable": 7783.93,
    "total": 9185.04,
    "docket": "20002075668",
    "transporter": "XP INDIA",
    "edd": "2026-11-01",
    "bookingDate": "2026-10-26"
  },
  {
    "id": "MVS/26-27/30896",
    "customer": "Carrier Refrigeration",
    "taxable": 5675.59,
    "total": 6697.2,
    "docket": "20002075676",
    "transporter": "XP INDIA",
    "edd": "2026-11-01",
    "bookingDate": "2026-10-26"
  },
  {
    "id": "MVS/26-27/30897",
    "customer": "Carrier Refrigeration",
    "taxable": 2880.61,
    "total": 3399.12,
    "docket": "20002075350",
    "transporter": "XP INDIA",
    "edd": "2026-11-01",
    "bookingDate": "2026-10-26"
  },
  {
    "id": "MVS/26-27/30899",
    "customer": "Carrier Refrigeration",
    "taxable": 4097.87,
    "total": 4835.49,
    "docket": "20002075352",
    "transporter": "XP INDIA",
    "edd": "2026-11-01",
    "bookingDate": "2026-10-26"
  },
  {
    "id": "MVS/26-27/30900",
    "customer": "Carrier Refrigeration",
    "taxable": 2091.63,
    "total": 2468.12,
    "docket": "20002075353",
    "transporter": "XP INDIA",
    "edd": "2026-11-01",
    "bookingDate": "2026-10-26"
  },
  {
    "id": "MVS/26-27/30901",
    "customer": "Carrier Refrigeration",
    "taxable": 2091.63,
    "total": 2468.12,
    "docket": "20002075354",
    "transporter": "XP INDIA",
    "edd": "2026-11-01",
    "bookingDate": "2026-10-26"
  },
  {
    "id": "MVS/26-27/30902",
    "customer": "Carrier Refrigeration",
    "taxable": 2122.81,
    "total": 2504.92,
    "docket": "20002075355",
    "transporter": "XP INDIA",
    "edd": "2026-11-01",
    "bookingDate": "2026-10-26"
  },
  {
    "id": "MVS/26-27/30904",
    "customer": "Carrier Refrigeration",
    "taxable": 5908.15,
    "total": 6971.62,
    "docket": "20002075357",
    "transporter": "XP INDIA",
    "edd": "2026-11-01",
    "bookingDate": "2026-10-26"
  },
  {
    "id": "MVS/26-27/30905",
    "customer": "Carrier Refrigeration",
    "taxable": 3105.14,
    "total": 3664.07,
    "docket": "20002075360",
    "transporter": "XP INDIA",
    "edd": "2026-11-01",
    "bookingDate": "2026-10-26"
  },
  {
    "id": "MVS/26-27/30906",
    "customer": "Carrier Refrigeration",
    "taxable": 6935.7,
    "total": 8184.13,
    "docket": "20002075361",
    "transporter": "XP INDIA",
    "edd": "2026-11-01",
    "bookingDate": "2026-10-26"
  },
  {
    "id": "MVS/26-27/30909",
    "customer": "Carrier Refrigeration",
    "taxable": 2357.74,
    "total": 2782.13,
    "docket": "20002075364",
    "transporter": "XP INDIA",
    "edd": "2026-11-01",
    "bookingDate": "2026-10-26"
  },
  {
    "id": "MVS/26-27/30913",
    "customer": "Carrier Refrigeration",
    "taxable": 334.88,
    "total": 395.16,
    "docket": "20002075368",
    "transporter": "XP INDIA",
    "edd": "2026-11-01",
    "bookingDate": "2026-10-26"
  },
  {
    "id": "MVS/26-27/30914",
    "customer": "Carrier Refrigeration",
    "taxable": 334.88,
    "total": 395.16,
    "docket": "20002075369",
    "transporter": "XP INDIA",
    "edd": "2026-11-01",
    "bookingDate": "2026-10-26"
  },
  {
    "id": "MVS/26-27/30916",
    "customer": "Carrier Refrigeration",
    "taxable": 334.88,
    "total": 395.16,
    "docket": "20002076049",
    "transporter": "XP INDIA",
    "edd": "2026-11-01",
    "bookingDate": "2026-10-26"
  },
  {
    "id": "MVS/26-27/30919",
    "customer": "Carrier Refrigeration",
    "taxable": 2076.04,
    "total": 2449.73,
    "docket": "20002076052",
    "transporter": "XP INDIA",
    "edd": "2026-11-01",
    "bookingDate": "2026-10-26"
  },
  {
    "id": "MVS/26-27/30920",
    "customer": "Carrier Refrigeration",
    "taxable": 3713.25,
    "total": 4381.64,
    "docket": "20002076231",
    "transporter": "XP INDIA",
    "edd": "2026-11-01",
    "bookingDate": "2026-10-26"
  },
  {
    "id": "MVS/26-27/30922",
    "customer": "Carrier Refrigeration",
    "taxable": 1729.89,
    "total": 2041.27,
    "docket": "20002076233",
    "transporter": "XP INDIA",
    "edd": "2026-11-01",
    "bookingDate": "2026-10-26"
  },
  {
    "id": "MVS/26-27/30926",
    "customer": "Carrier Refrigeration",
    "taxable": 7415.95,
    "total": 8750.82,
    "docket": "20002076237",
    "transporter": "XP INDIA",
    "edd": "2026-11-01",
    "bookingDate": "2026-10-26"
  },
  {
    "id": "MVS/26-27/30928",
    "customer": "Carrier Refrigeration",
    "taxable": 8337.22,
    "total": 9837.92,
    "docket": "20002076372",
    "transporter": "XP INDIA",
    "edd": "2026-11-01",
    "bookingDate": "2026-10-26"
  },
  {
    "id": "MVS/26-27/30938",
    "customer": "Sunrise Sports India Pvt Ltd",
    "taxable": 2179.86,
    "total": 2572.23,
    "docket": "20002077956",
    "transporter": "XP INDIA",
    "edd": "2026-11-01",
    "bookingDate": "2026-10-27"
  },
  {
    "id": "MVS/26-27/30944",
    "customer": "Sunrise Sports India Pvt Ltd",
    "taxable": 1499.81,
    "total": 1769.78,
    "docket": "20002079081",
    "transporter": "XP INDIA",
    "edd": "2026-11-01",
    "bookingDate": "2026-10-27"
  },
  {
    "id": "MVS/26-27/30945",
    "customer": "Sunrise Sports India Pvt Ltd",
    "taxable": 15114.26,
    "total": 17834.83,
    "docket": "20002079044",
    "transporter": "XP INDIA",
    "edd": "2026-11-01",
    "bookingDate": "2026-10-27"
  },
  {
    "id": "MVS/26-27/30946",
    "customer": "Sunrise Sports India Pvt Ltd",
    "taxable": 590.16,
    "total": 696.39,
    "docket": "20002079237",
    "transporter": "XP INDIA",
    "edd": "2026-11-01",
    "bookingDate": "2026-10-27"
  },
  {
    "id": "MVS/26-27/30954",
    "customer": "Carrier Refrigeration",
    "taxable": 2406.6,
    "total": 2839.79,
    "docket": "540261012390",
    "transporter": "OM",
    "edd": "2026-11-01",
    "bookingDate": "2026-10-27"
  },
  {
    "id": "MVS/26-27/30958",
    "customer": "Carrier - CTD",
    "taxable": 8685.98,
    "total": 10249.46,
    "docket": "1303726750",
    "transporter": "DP WORLD",
    "edd": "2026-11-01",
    "bookingDate": "2026-10-27"
  },
  {
    "id": "MVS/26-27/30961",
    "customer": "Haier CCR",
    "taxable": 3176.02,
    "total": 3747.7,
    "docket": "20002079343",
    "transporter": "XP INDIA",
    "edd": "2026-11-01",
    "bookingDate": "2026-10-27"
  },
  {
    "id": "MVS/26-27/30967",
    "customer": "Oneric Appliances",
    "taxable": 1468.8,
    "total": 1733.18,
    "docket": "1341355521",
    "transporter": "DP WORLD",
    "edd": "2026-11-01",
    "bookingDate": "2026-10-27"
  },
  {
    "id": "MVS/26-27/30976",
    "customer": "Bombax",
    "taxable": 443.75,
    "total": 523.63,
    "docket": "2016885669",
    "transporter": "EKART",
    "edd": "2026-11-01",
    "bookingDate": "2026-10-27"
  },
  {
    "id": "MVS/26-27/30977",
    "customer": "Bombax",
    "taxable": 443.75,
    "total": 523.63,
    "docket": "2016883671",
    "transporter": "EKART",
    "edd": "2026-11-01",
    "bookingDate": "2026-10-27"
  },
  {
    "id": "MVS/26-27/30978",
    "customer": "Bombax",
    "taxable": 622.35,
    "total": 734.37,
    "docket": "2016881673",
    "transporter": "EKART",
    "edd": "2026-11-01",
    "bookingDate": "2026-10-27"
  },
  {
    "id": "MVS/26-27/30980",
    "customer": "Bombax",
    "taxable": 622.35,
    "total": 734.37,
    "docket": "2016889681",
    "transporter": "EKART",
    "edd": "2026-11-01",
    "bookingDate": "2026-10-27"
  },
  {
    "id": "MVS/26-27/30997",
    "customer": "Carrier - CTD",
    "taxable": 8890.97,
    "total": 10491.34,
    "docket": "1303726751",
    "transporter": "DP WORLD",
    "edd": "2026-11-01",
    "bookingDate": "2026-10-27"
  },
  {
    "id": "MVS/26-27/30998",
    "customer": "Carrier - CTD",
    "taxable": 5953.15,
    "total": 7024.72,
    "docket": "1303726752",
    "transporter": "DP WORLD",
    "edd": "2026-11-01",
    "bookingDate": "2026-10-27"
  },
  {
    "id": "MVS/26-27/30999",
    "customer": "Carrier - CTD",
    "taxable": 3194.54,
    "total": 3769.56,
    "docket": "1303726753",
    "transporter": "DP WORLD",
    "edd": "2026-11-01",
    "bookingDate": "2026-10-27"
  },
  {
    "id": "MVS/26-27/31000",
    "customer": "Carrier - CTD",
    "taxable": 8685.98,
    "total": 10249.46,
    "docket": "1303726754",
    "transporter": "DP WORLD",
    "edd": "2026-11-01",
    "bookingDate": "2026-10-27"
  },
  {
    "id": "MVS/26-27/31001",
    "customer": "Carrier - CTD",
    "taxable": 3379.22,
    "total": 3987.47,
    "docket": "1303726755",
    "transporter": "DP WORLD",
    "edd": "2026-11-01",
    "bookingDate": "2026-10-27"
  },
  {
    "id": "MVS/26-27/31002",
    "customer": "Carrier Refrigeration",
    "taxable": 334.88,
    "total": 395.16,
    "docket": "20002075963",
    "transporter": "XP INDIA",
    "edd": "2026-11-01",
    "bookingDate": "2026-10-27"
  },
  {
    "id": "MVS/26-27/31003",
    "customer": "Carrier Refrigeration",
    "taxable": 1824.48,
    "total": 2152.89,
    "docket": "20002077889",
    "transporter": "XP INDIA",
    "edd": "2026-11-01",
    "bookingDate": "2026-10-27"
  },
  {
    "id": "MVS/26-27/31006",
    "customer": "Carrier Refrigeration",
    "taxable": 2293.29,
    "total": 2706.08,
    "docket": "20002077895",
    "transporter": "XP INDIA",
    "edd": "2026-11-01",
    "bookingDate": "2026-10-27"
  },
  {
    "id": "MVS/26-27/31007",
    "customer": "Carrier Refrigeration",
    "taxable": 8131.81,
    "total": 9595.54,
    "docket": "20002077891",
    "transporter": "XP INDIA",
    "edd": "2026-11-01",
    "bookingDate": "2026-10-27"
  },
  {
    "id": "MVS/26-27/31008",
    "customer": "Carrier Refrigeration",
    "taxable": 1702.86,
    "total": 2009.37,
    "docket": "20002079458",
    "transporter": "XP INDIA",
    "edd": "2026-11-01",
    "bookingDate": "2026-10-27"
  },
  {
    "id": "MVS/26-27/31009",
    "customer": "Carrier Refrigeration",
    "taxable": 2018.86,
    "total": 2382.25,
    "docket": "20002077893",
    "transporter": "XP INDIA",
    "edd": "2026-11-01",
    "bookingDate": "2026-10-27"
  },
  {
    "id": "MVS/26-27/31010",
    "customer": "Carrier Refrigeration",
    "taxable": 10755.14,
    "total": 12691.07,
    "docket": "20002077892",
    "transporter": "XP INDIA",
    "edd": "2026-11-01",
    "bookingDate": "2026-10-27"
  },
  {
    "id": "MVS/26-27/31014",
    "customer": "Carrier Refrigeration",
    "taxable": 932.59,
    "total": 1100.46,
    "docket": "20002078319",
    "transporter": "XP INDIA",
    "edd": "2026-11-01",
    "bookingDate": "2026-10-27"
  },
  {
    "id": "MVS/26-27/31015",
    "customer": "Carrier Refrigeration",
    "taxable": 334.88,
    "total": 395.16,
    "docket": "20002078315",
    "transporter": "XP INDIA",
    "edd": "2026-11-01",
    "bookingDate": "2026-10-27"
  },
  {
    "id": "MVS/26-27/31017",
    "customer": "Carrier Refrigeration",
    "taxable": 937.78,
    "total": 1106.58,
    "docket": "20002078320",
    "transporter": "XP INDIA",
    "edd": "2026-11-01",
    "bookingDate": "2026-10-27"
  },
  {
    "id": "MVS/26-27/31020",
    "customer": "Carrier Refrigeration",
    "taxable": 334.88,
    "total": 395.16,
    "docket": "20002078318",
    "transporter": "XP INDIA",
    "edd": "2026-11-01",
    "bookingDate": "2026-10-27"
  },
  {
    "id": "MVS/26-27/31021",
    "customer": "Carrier Refrigeration",
    "taxable": 1607.22,
    "total": 1896.52,
    "docket": "20002078322",
    "transporter": "XP INDIA",
    "edd": "2026-11-01",
    "bookingDate": "2026-10-27"
  },
  {
    "id": "MVS/26-27/31022",
    "customer": "Carrier Refrigeration",
    "taxable": 750.68,
    "total": 885.8,
    "docket": "20002078313",
    "transporter": "XP INDIA",
    "edd": "2026-11-01",
    "bookingDate": "2026-10-27"
  },
  {
    "id": "MVS/26-27/31025",
    "customer": "Carrier Refrigeration",
    "taxable": 15866.04,
    "total": 18721.93,
    "docket": "20002078327",
    "transporter": "XP INDIA",
    "edd": "2026-11-01",
    "bookingDate": "2026-10-27"
  },
  {
    "id": "MVS/26-27/31026",
    "customer": "Carrier Refrigeration",
    "taxable": 334.88,
    "total": 395.16,
    "docket": "20002078316",
    "transporter": "XP INDIA",
    "edd": "2026-11-01",
    "bookingDate": "2026-10-27"
  },
  {
    "id": "MVS/26-27/31027",
    "customer": "Carrier Refrigeration",
    "taxable": 635.29,
    "total": 749.64,
    "docket": "20002078330",
    "transporter": "XP INDIA",
    "edd": "2026-11-01",
    "bookingDate": "2026-10-27"
  },
  {
    "id": "MVS/26-27/31030",
    "customer": "Carrier Refrigeration",
    "taxable": 1332.8,
    "total": 1572.7,
    "docket": "20002078323",
    "transporter": "XP INDIA",
    "edd": "2026-11-01",
    "bookingDate": "2026-10-27"
  },
  {
    "id": "MVS/26-27/31039",
    "customer": "Carrier Refrigeration",
    "taxable": 1885.81,
    "total": 2225.26,
    "docket": "2016884822",
    "transporter": "EKART",
    "edd": "2026-11-01",
    "bookingDate": "2026-10-27"
  },
  {
    "id": "MVS/26-27/31041",
    "customer": "Carrier Refrigeration",
    "taxable": 2764.19,
    "total": 3261.74,
    "docket": "2016889844",
    "transporter": "EKART",
    "edd": "2026-11-01",
    "bookingDate": "2026-10-27"
  },
  {
    "id": "MVS/26-27/31042",
    "customer": "Carrier Refrigeration",
    "taxable": 2291.22,
    "total": 2703.64,
    "docket": "2016887861",
    "transporter": "EKART",
    "edd": "2026-11-01",
    "bookingDate": "2026-10-27"
  },
  {
    "id": "MVS/26-27/31046",
    "customer": "Carrier Refrigeration",
    "taxable": 334.88,
    "total": 395.16,
    "docket": "20002078615",
    "transporter": "XP INDIA",
    "edd": "2026-11-01",
    "bookingDate": "2026-10-27"
  },
  {
    "id": "MVS/26-27/31048",
    "customer": "Carrier Refrigeration",
    "taxable": 3567.72,
    "total": 4209.91,
    "docket": "20002078617",
    "transporter": "XP INDIA",
    "edd": "2026-11-01",
    "bookingDate": "2026-10-27"
  },
  {
    "id": "MVS/26-27/31052",
    "customer": "Carrier Refrigeration",
    "taxable": 1693.5,
    "total": 1998.33,
    "docket": "20002078621",
    "transporter": "XP INDIA",
    "edd": "2026-11-01",
    "bookingDate": "2026-10-27"
  },
  {
    "id": "MVS/26-27/31054",
    "customer": "Carrier Refrigeration",
    "taxable": 334.88,
    "total": 395.16,
    "docket": "20002078828",
    "transporter": "XP INDIA",
    "edd": "2026-11-01",
    "bookingDate": "2026-10-27"
  },
  {
    "id": "MVS/26-27/31060",
    "customer": "Carrier Refrigeration",
    "taxable": 2533.42,
    "total": 2989.44,
    "docket": "20002078835",
    "transporter": "XP INDIA",
    "edd": "2026-11-01",
    "bookingDate": "2026-10-27"
  },
  {
    "id": "MVS/26-27/31065",
    "customer": "Carrier Refrigeration",
    "taxable": 1864.8,
    "total": 2200.46,
    "docket": "1303645149",
    "transporter": "DP WORLD",
    "edd": "2026-11-01",
    "bookingDate": "2026-10-27"
  },
  {
    "id": "MVS/26-27/31077",
    "customer": "Carrier Refrigeration",
    "taxable": 6178.59,
    "total": 7290.74,
    "docket": "1303387276",
    "transporter": "DP WORLD",
    "edd": "2026-11-01",
    "bookingDate": "2026-10-27"
  },
  {
    "id": "MVS/26-27/31086",
    "customer": "Loom Solar",
    "taxable": 2968.0,
    "total": 3502.24,
    "docket": "2016902686",
    "transporter": "EKART",
    "edd": "2026-11-01",
    "bookingDate": "2026-10-28"
  },
  {
    "id": "MVS/26-27/31090",
    "customer": "MITRAS",
    "taxable": 5804.08,
    "total": 6848.81,
    "docket": "1303686131",
    "transporter": "DP WORLD",
    "edd": "2026-11-01",
    "bookingDate": "2026-10-28"
  },
  {
    "id": "MVS/26-27/31091",
    "customer": "Sunrise Sports India Pvt Ltd",
    "taxable": 557.3,
    "total": 657.61,
    "docket": "1341355524",
    "transporter": "DP WORLD",
    "edd": "2026-11-01",
    "bookingDate": "2026-10-28"
  },
  {
    "id": "MVS/26-27/31092",
    "customer": "Sunrise Sports India Pvt Ltd",
    "taxable": 526.69,
    "total": 621.49,
    "docket": "1341355525",
    "transporter": "DP WORLD",
    "edd": "2026-11-01",
    "bookingDate": "2026-10-28"
  },
  {
    "id": "MVS/26-27/31093",
    "customer": "Sunrise Sports India Pvt Ltd",
    "taxable": 557.3,
    "total": 657.61,
    "docket": "1341355526",
    "transporter": "DP WORLD",
    "edd": "2026-11-01",
    "bookingDate": "2026-10-28"
  },
  {
    "id": "MVS/26-27/31094",
    "customer": "Kumar Services",
    "taxable": 2516.32,
    "total": 2969.26,
    "docket": "1303594431",
    "transporter": "DP WORLD",
    "edd": "2026-11-01",
    "bookingDate": "2026-10-28"
  },
  {
    "id": "MVS/26-27/31097",
    "customer": "Carrier Refrigeration",
    "taxable": 1799.41,
    "total": 2123.3,
    "docket": "1303589030",
    "transporter": "DP WORLD",
    "edd": "2026-11-01",
    "bookingDate": "2026-10-28"
  },
  {
    "id": "MVS/26-27/31102",
    "customer": "Carrier Refrigeration",
    "taxable": 2534.87,
    "total": 2991.15,
    "docket": "1303589037",
    "transporter": "DP WORLD",
    "edd": "2026-11-01",
    "bookingDate": "2026-10-28"
  },
  {
    "id": "MVS/26-27/31107",
    "customer": "Sunrise Sports India Pvt Ltd",
    "taxable": 2373.12,
    "total": 2800.28,
    "docket": "20002080989",
    "transporter": "XP INDIA",
    "edd": "2026-11-01",
    "bookingDate": "2026-10-28"
  },
  {
    "id": "MVS/26-27/31111",
    "customer": "Carrier Refrigeration",
    "taxable": 2231.09,
    "total": 2632.69,
    "docket": "9000235477",
    "transporter": "EKART",
    "edd": "2026-11-01",
    "bookingDate": "2026-10-28"
  },
  {
    "id": "MVS/26-27/31113",
    "customer": "Haier CCR",
    "taxable": 2788.8,
    "total": 3290.78,
    "docket": "1303557990",
    "transporter": "DP WORLD",
    "edd": "2026-11-01",
    "bookingDate": "2026-10-28"
  },
  {
    "id": "MVS/26-27/31116",
    "customer": "Carrier Refrigeration",
    "taxable": 10679.27,
    "total": 12601.54,
    "docket": "1303487500",
    "transporter": "DP WORLD",
    "edd": "2026-11-01",
    "bookingDate": "2026-10-28"
  },
  {
    "id": "MVS/26-27/31156",
    "customer": "Sukuga",
    "taxable": 1112.11,
    "total": 1312.3,
    "docket": "1303129506",
    "transporter": "DP WORLD",
    "edd": "2026-11-01",
    "bookingDate": "2026-10-28"
  },
  {
    "id": "MVS/26-27/31176",
    "customer": "Carrier Refrigeration",
    "taxable": 8195.57,
    "total": 9670.77,
    "docket": "20002080788",
    "transporter": "XP INDIA",
    "edd": "2026-11-01",
    "bookingDate": "2026-10-28"
  },
  {
    "id": "MVS/26-27/31186",
    "customer": "Carrier Refrigeration",
    "taxable": 2568.24,
    "total": 3030.52,
    "docket": "2016918740",
    "transporter": "EKART",
    "edd": "2026-11-01",
    "bookingDate": "2026-10-28"
  },
  {
    "id": "MVS/26-27/31213",
    "customer": "Carrier Refrigeration",
    "taxable": 334.88,
    "total": 395.16,
    "docket": "20002081098",
    "transporter": "XP INDIA",
    "edd": "2026-11-01",
    "bookingDate": "2026-10-28"
  },
  {
    "id": "MVS/26-27/31225",
    "customer": "Carrier Refrigeration",
    "taxable": 334.88,
    "total": 395.16,
    "docket": "20002082808",
    "transporter": "XP INDIA",
    "edd": "2026-11-01",
    "bookingDate": "2026-10-28"
  },
  {
    "id": "MVS/26-27/31243",
    "customer": "Carrier Refrigeration",
    "taxable": 334.88,
    "total": 395.16,
    "docket": "20002081815",
    "transporter": "XP INDIA",
    "edd": "2026-11-01",
    "bookingDate": "2026-10-28"
  },
  {
    "id": "MVS/26-27/31249",
    "customer": "Carrier Refrigeration",
    "taxable": 1261.07,
    "total": 1488.06,
    "docket": "20002081808",
    "transporter": "XP INDIA",
    "edd": "2026-11-01",
    "bookingDate": "2026-10-28"
  },
  {
    "id": "MVS/26-27/31252",
    "customer": "Carrier Refrigeration",
    "taxable": 693.5,
    "total": 818.33,
    "docket": "20002081805",
    "transporter": "XP INDIA",
    "edd": "2026-11-01",
    "bookingDate": "2026-10-28"
  },
  {
    "id": "MVS/26-27/31263",
    "customer": "Carrier Refrigeration",
    "taxable": 838.24,
    "total": 989.12,
    "docket": "1303726850",
    "transporter": "DP WORLD",
    "edd": "2026-11-01",
    "bookingDate": "2026-10-28"
  },
  {
    "id": "MVS/26-27/31274",
    "customer": "Carrier Refrigeration",
    "taxable": 13644.8,
    "total": 16100.86,
    "docket": "1303680236",
    "transporter": "DP WORLD",
    "edd": "2026-11-01",
    "bookingDate": "2026-10-28"
  },
  {
    "id": "MVS/26-27/31275",
    "customer": "Carrier Refrigeration",
    "taxable": 2374.37,
    "total": 2801.76,
    "docket": "1303680238",
    "transporter": "DP WORLD",
    "edd": "2026-11-01",
    "bookingDate": "2026-10-28"
  },
  {
    "id": "MVS/26-27/31284",
    "customer": "Carrier Refrigeration",
    "taxable": 1516.41,
    "total": 1789.37,
    "docket": "1303387280",
    "transporter": "DP WORLD",
    "edd": "2026-11-01",
    "bookingDate": "2026-10-28"
  },
  {
    "id": "MVS/26-27/31309",
    "customer": "Carrier - CTD",
    "taxable": 334.88,
    "total": 395.16,
    "docket": "1303726689",
    "transporter": "DP WORLD",
    "edd": "2026-11-01",
    "bookingDate": "2026-10-29"
  },
  {
    "id": "MVS/26-27/31323",
    "customer": "Sukuga",
    "taxable": 5238.8,
    "total": 6181.78,
    "docket": "1303367231",
    "transporter": "DP WORLD",
    "edd": "2026-11-01",
    "bookingDate": "2026-10-29"
  },
  {
    "id": "MVS/26-27/31334",
    "customer": "Indu Sports",
    "taxable": 2131.8,
    "total": 2515.52,
    "docket": "1341355535",
    "transporter": "DP WORLD",
    "edd": "2026-11-01",
    "bookingDate": "2026-10-29"
  },
  {
    "id": "MVS/26-27/31338",
    "customer": "Kumar Services",
    "taxable": 12507.28,
    "total": 14758.59,
    "docket": "1341355529",
    "transporter": "DP WORLD",
    "edd": "2026-11-01",
    "bookingDate": "2026-10-29"
  },
  {
    "id": "MVS/26-27/31340",
    "customer": "Carrier Refrigeration",
    "taxable": 549.01,
    "total": 647.83,
    "docket": "1303589024",
    "transporter": "DP WORLD",
    "edd": "2026-11-01",
    "bookingDate": "2026-10-29"
  },
  {
    "id": "MVS/26-27/31367",
    "customer": "Indu Sports",
    "taxable": 2131.8,
    "total": 2515.52,
    "docket": "1341355539",
    "transporter": "DP WORLD",
    "edd": "2026-11-01",
    "bookingDate": "2026-10-29"
  },
  {
    "id": "MVS/26-27/31388",
    "customer": "Carrier Refrigeration",
    "taxable": 5189.86,
    "total": 6124.03,
    "docket": "20002084368",
    "transporter": "XP INDIA",
    "edd": "2026-11-01",
    "bookingDate": "2026-10-29"
  },
  {
    "id": "MVS/26-27/31392",
    "customer": "Carrier Refrigeration",
    "taxable": 493.92,
    "total": 582.83,
    "docket": "20002084372",
    "transporter": "XP INDIA",
    "edd": "2026-11-01",
    "bookingDate": "2026-10-29"
  },
  {
    "id": "MVS/26-27/31415",
    "customer": "Carrier Refrigeration",
    "taxable": 831.76,
    "total": 981.48,
    "docket": "2016923813",
    "transporter": "EKART",
    "edd": "2026-11-01",
    "bookingDate": "2026-10-29"
  },
  {
    "id": "MVS/26-27/31440",
    "customer": "Carrier Refrigeration",
    "taxable": 439.86,
    "total": 519.03,
    "docket": "20002084564",
    "transporter": "XP INDIA",
    "edd": "2026-11-01",
    "bookingDate": "2026-10-29"
  },
  {
    "id": "MVS/26-27/31491",
    "customer": "Carrier Refrigeration",
    "taxable": 617.03,
    "total": 728.1,
    "docket": "1303349022",
    "transporter": "DP WORLD",
    "edd": "2026-11-01",
    "bookingDate": "2026-10-29"
  },
  {
    "id": "MVS/26-27/31493",
    "customer": "Carrier Refrigeration",
    "taxable": 1354.26,
    "total": 1598.02,
    "docket": "1303387825",
    "transporter": "DP WORLD",
    "edd": "2026-11-01",
    "bookingDate": "2026-10-29"
  },
  {
    "id": "MVS/26-27/31607",
    "customer": "KTB",
    "taxable": 3850.0,
    "total": 4543.0,
    "docket": "NULL",
    "transporter": "NULL",
    "edd": "2026-11-01",
    "bookingDate": "2026-10-30"
  }
];

// 4. Open Shipments counts by customer (948 orders from Open Shipment sheet)
const openCounts = {
  "Bombax": 142,
  "Khusbhu Enterprises": 2,
  "MVIKAS Tech": 5,
  "Carrier Refrigeration": 549,
  "Abhishek Tomar": 6,
  "Haier CCR": 24,
  "Sunrise Sports India Pvt Ltd": 92,
  "MS Medicals Pvt Ltd": 4,
  "Khatana Enterprises": 7,
  "Cosmos Pumps Pvt. Ltd.": 6,
  "Loom Solar": 10,
  "Droplets Engineering": 11,
  "Carrier - CTD": 62,
  "Vaidrishi Laboratories": 3,
  "Oneric Appliances": 4,
  "MITRAS": 1,
  "Kumar Services": 2,
  "Sukuga": 9,
  "Indu Sports": 8,
  "KTB": 1
};
const openTotal = Object.values(openCounts).reduce((a, b) => a + b, 0);

// 5. Booked Yesterday by Customer (239 orders from Shipment Booked Yesterday sheet)
const bookedCustList = [
  {
    "name": "Carrier Refrigeration",
    "count": 123
  },
  {
    "name": "Bombax",
    "count": 50
  },
  {
    "name": "Carrier - CTD",
    "count": 22
  },
  {
    "name": "Sunrise Sports India Pvt Ltd",
    "count": 14
  },
  {
    "name": "Haier CCR",
    "count": 9
  },
  {
    "name": "Kumar Services",
    "count": 5
  },
  {
    "name": "Loom Solar",
    "count": 4
  },
  {
    "name": "Sukuga",
    "count": 4
  },
  {
    "name": "Oneric Appliances",
    "count": 4
  },
  {
    "name": "Cosmos Pumps Pvt. Ltd.",
    "count": 2
  },
  {
    "name": "Vaidrishi Laboratories",
    "count": 1
  }
];

// Aggregates
const dueTotal = dueOrders.length;
const inTransitTotal = Math.max(openTotal - dueTotal, 0);
const bookedTotal = 239; // from Shipment Booked Yesterday sheet

const monthlyTotalTonnage = +(clients.reduce((a, c) => a + c.achieved, 0)).toFixed(2);
const monthlyTargetTonnage = +(clients.reduce((a, c) => a + c.target, 0)).toFixed(2);
const activeDays = 20;
const dailyAverage = +(monthlyTotalTonnage / activeDays).toFixed(2);

// Finance Data directly from sheets (NO average rate multipliers!)
const totalTaxableAchieved = clients.reduce((s, c) => s + (c.taxable || 0), 0);
const totalFinancialTarget = clients.reduce((s, c) => s + (c.finTarget || 0), 0);

// Currency and Weight Formatters (T only, no kg)
const fT = v => Number(v || 0).toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + ' T';
const fTUnit = v => fT(v);
const fTVal = v => Number(v || 0).toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
const fR = v => '₹' + Math.round(v || 0).toLocaleString('en-IN');
const fRDec = v => '₹' + Number(v || 0).toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
const fLakh = v => '₹' + (v / 100000).toFixed(2) + 'L';

// Populate KPIs
if ($('kpi-open')) $('kpi-open').textContent = openTotal.toLocaleString('en-IN');
if ($('kpi-daily-avg')) $('kpi-daily-avg').innerHTML = fTVal(dailyAverage) + ' <span style="font-size:.75rem;font-weight:500;color:' + C.text3 + '">T/day</span>';
if ($('kpi-due')) $('kpi-due').textContent = dueTotal;
if ($('kpi-due-pct')) $('kpi-due-pct').textContent = (dueTotal / openTotal * 100).toFixed(1) + '% of open';
if ($('kpi-booked')) $('kpi-booked').textContent = bookedTotal;
if ($('kpi-daily-ton')) $('kpi-daily-ton').innerHTML = fTVal(dailyTotalTonnage) + ' <span style="font-size:.75rem;font-weight:500;color:' + C.text3 + '">T</span>';
if ($('kpi-month-ton')) $('kpi-month-ton').innerHTML = fTVal(monthlyTotalTonnage) + ' <span style="font-size:.75rem;font-weight:500;color:' + C.text3 + '">T</span>';
if ($('kpi-target-finance')) $('kpi-target-finance').textContent = '₹' + (totalFinancialTarget / 100000).toFixed(2) + ' Lakhs';
if ($('kpi-sales-finance')) $('kpi-sales-finance').textContent = '₹' + (totalTaxableAchieved / 100000).toFixed(2) + ' Lakhs';

// Hero Stats
const heroOpen = $('hero-open');
if (heroOpen) { heroOpen.dataset.target = openTotal; heroOpen.textContent = openTotal; }
const heroDue = $('hero-due');
if (heroDue) { heroDue.dataset.target = dueTotal; heroDue.textContent = dueTotal; }
const heroTon = $('hero-ton');
if (heroTon) { heroTon.dataset.target = monthlyTotalTonnage.toFixed(2); heroTon.textContent = monthlyTotalTonnage.toFixed(2); }
const heroFinance = $('hero-finance');
if (heroFinance) {
  const finLakhs = (totalTaxableAchieved / 100000).toFixed(2);
  heroFinance.dataset.target = finLakhs;
  heroFinance.textContent = finLakhs;
}

// Strip Values
if ($('strip-transit')) $('strip-transit').textContent = inTransitTotal.toLocaleString('en-IN');
if ($('strip-due')) $('strip-due').textContent = dueTotal.toLocaleString('en-IN');
if ($('strip-booked')) $('strip-booked').textContent = bookedTotal.toLocaleString('en-IN');
if ($('strip-total')) $('strip-total').textContent = openTotal.toLocaleString('en-IN');

if ($('dl-transit')) $('dl-transit').textContent = inTransitTotal.toLocaleString('en-IN');
if ($('dl-due')) $('dl-due').textContent = dueTotal.toLocaleString('en-IN');
if ($('dl-booked')) $('dl-booked').textContent = bookedTotal.toLocaleString('en-IN');
if ($('dl-total')) $('dl-total').textContent = openTotal.toLocaleString('en-IN');

// ═══════════════════════════════════════════════════════════════
// CHARTS — PANEL 1 (CLASSIFICATION & DUE TOMORROW)
// ═══════════════════════════════════════════════════════════════

// 1. Status Donut Chart
safeChart($('statusDonut'), {
  type: 'doughnut',
  data: {
    labels: ['In Transit', 'Due Tomorrow', 'Booked Yesterday'],
    datasets: [{
      data: [inTransitTotal, dueTotal, bookedTotal],
      backgroundColor: [C.orange, C.yellow, C.green],
      borderColor: C.bg,
      borderWidth: 3,
      hoverOffset: 6
    }]
  },
  options: {
    responsive: true,
    maintainAspectRatio: false,
    cutout: '72%',
    plugins: {
      legend: { display: false },
      tooltip: {
        callbacks: {
          label: ctx => ` ${ctx.label}: ${ctx.raw} orders (${(ctx.raw / openTotal * 100).toFixed(1)}%)`
        }
      }
    }
  }
});

// Group due orders by customer
const dueByCustMap = {};
dueOrders.forEach(o => {
  if (!dueByCustMap[o.customer]) {
    dueByCustMap[o.customer] = { count: 0, taxable: 0, total: 0 };
  }
  dueByCustMap[o.customer].count += 1;
  dueByCustMap[o.customer].taxable += o.taxable;
  dueByCustMap[o.customer].total += o.total;
});
const dueCustList = Object.entries(dueByCustMap)
  .map(([name, d]) => ({ name, ...d }))
  .sort((a, b) => b.count - a.count);

// 2. Due Tomorrow Orders by Customer
safeChart($('dueTmrChart'), {
  type: 'bar',
  data: {
    labels: dueCustList.map(c => c.name),
    datasets: [{
      label: 'Due Orders',
      data: dueCustList.map(c => c.count),
      backgroundColor: orangeA(0.85),
      borderRadius: 4
    }]
  },
  options: {
    indexAxis: 'y',
    responsive: true,
    maintainAspectRatio: false,
    plugins: {
      legend: { display: false },
      tooltip: {
        callbacks: {
          label: ctx => ` Orders Due: ${ctx.raw}`
        }
      }
    }
  }
});

// Group due orders by transporter
const dueTransMap = {};
dueOrders.forEach(o => {
  const t = o.transporter || 'Direct/Other';
  if (!dueTransMap[t]) {
    dueTransMap[t] = { count: 0, taxable: 0, total: 0 };
  }
  dueTransMap[t].count += 1;
  dueTransMap[t].taxable += o.taxable;
  dueTransMap[t].total += o.total;
});
const dueTransList = Object.entries(dueTransMap)
  .map(([name, d]) => ({ name, ...d }))
  .sort((a, b) => b.count - a.count);

// 3. Due Tomorrow by Transporter
safeChart($('transporterChart'), {
  type: 'bar',
  data: {
    labels: dueTransList.map(t => t.name),
    datasets: [{
      label: 'Orders',
      data: dueTransList.map(t => t.count),
      backgroundColor: [C.orange, '#8b8b8b', '#5a5a5a', '#3a3a3a', '#2a2a2a'],
      borderRadius: 4
    }]
  },
  options: {
    responsive: true,
    maintainAspectRatio: false,
    plugins: {
      legend: { display: false },
      tooltip: {
        callbacks: {
          label: ctx => ` ${ctx.label}: ${ctx.raw} orders`
        }
      }
    }
  }
});

// 4. Booked Yesterday by Customer (from Shipment Booked Yesterday sheet)
const topBookedCusts = bookedCustList.slice(0, 8);
safeChart($('bookedChart'), {
  type: 'bar',
  data: {
    labels: topBookedCusts.map(c => c.name),
    datasets: [{
      label: 'Booked Orders',
      data: topBookedCusts.map(c => c.count),
      backgroundColor: C.green,
      borderRadius: 4
    }]
  },
  options: {
    indexAxis: 'y',
    responsive: true,
    maintainAspectRatio: false,
    plugins: {
      legend: { display: false },
      tooltip: {
        callbacks: {
          label: ctx => ` Booked: ${ctx.raw} orders`
        }
      }
    }
  }
});

// ═══════════════════════════════════════════════════════════════
// PANEL 2: TONNAGE PROGRESS & KAM SUMMARY (IN TONNES T)
// ═══════════════════════════════════════════════════════════════
const PL = { full: 'Tonnage — % Completion (Full Month)', first10: 'First 10 Days', mid10: 'Mid 10 Days', last10: 'Last 10 Days' };
let curPeriod = 'full';

function getPStats(c, p) {
  if (p === 'full') return { target: c.target, achieved: c.achieved, hasTarget: c.target > 0, hasAchieved: c.achieved > 0 };
  const d = (c.periods || {})[p] || {};
  return { target: d.target || 0, achieved: d.achieved || 0, hasTarget: (d.target || 0) > 0, hasAchieved: (d.achieved || 0) > 0 };
}

function setTonnagePeriod(p) {
  curPeriod = p;
  renderTonnageBars();
  renderTonnageCharts();
}
window.setTonnagePeriod = setTonnagePeriod;

function renderTonnageBars() {
  const el = $('tonnage-bars');
  if (!el) return;
  const ti = $('tonnage-panel-title');
  if (ti) ti.textContent = PL[curPeriod] || 'Tonnage — % Completion';

  const rows = clients.map(c => {
    const s = getPStats(c, curPeriod);
    const pct = s.target > 0 ? Math.round(s.achieved / s.target * 100) : (s.achieved > 0 ? 100 : 0);
    return { ...c, pTarget: s.target, pAchieved: s.achieved, pPct: pct };
  }).sort((a, b) => b.pAchieved - a.pAchieved);

  el.innerHTML = rows.map(c => {
    const barWidth = Math.min(c.pPct, 100);
    const badgeColor = c.pPct >= 100 ? '#10b981' : (c.pPct >= 70 ? C.orange : C.text3);
    return `<div class="client-row">
      <div class="client-name" title="${c.name}">
        <span class="client-name-text">${c.name}</span>
      </div>
      <div class="client-person" title="KAM: ${c.person}">${c.person}</div>
      <div class="prog-bar-wrap">
        <div class="prog-bar" style="width:${barWidth}%;background:${badgeColor}"></div>
      </div>
      <div class="pct-text" style="color:${badgeColor}">${c.pPct}%</div>
      <div class="client-tonnage"><span style="color:var(--text);font-weight:600">${fTVal(c.pAchieved)}</span> <span style="color:var(--text-4)">/</span> <span style="color:var(--text-3)">${fT(c.pTarget)}</span></div>
    </div>`;
  }).join('');
}

let targetChartInst = null;
function renderTonnageCharts() {
  const top8 = [...clients].sort((a, b) => b.achieved - a.achieved).slice(0, 8);
  const ctx = $('targetChart');
  if (!ctx) return;
  if (targetChartInst) targetChartInst.destroy();

  targetChartInst = safeChart(ctx, {
    type: 'bar',
    data: {
      labels: top8.map(c => c.name),
      datasets: [
        {
          label: 'Target (T)',
          data: top8.map(c => c.target),
          backgroundColor: '#3a3a3a',
          borderRadius: 4
        },
        {
          label: 'Achieved (T)',
          data: top8.map(c => c.achieved),
          backgroundColor: C.orange,
          borderRadius: 4
        }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      scales: {
        y: {
          ticks: {
            callback: v => v + ' T'
          }
        }
      },
      plugins: {
        legend: { display: true, position: 'top', labels: { boxWidth: 12 } },
        tooltip: {
          callbacks: {
            label: ctx => ` ${ctx.dataset.label}: ${ctx.raw} T`
          }
        }
      }
    }
  });
}

function renderKamSummary() {
  const kams = {};
  clients.forEach(c => {
    const p = c.person || 'Direct';
    if (!kams[p]) kams[p] = { target: 0, achieved: 0, finTarget: 0, finAchieved: 0 };
    kams[p].target += c.target;
    kams[p].achieved += c.achieved;
    kams[p].finTarget += c.finTarget || 0;
    kams[p].finAchieved += c.taxable || 0;
  });

  const kamList = Object.entries(kams).map(([person, d]) => ({
    person,
    target: +d.target.toFixed(2),
    achieved: +d.achieved.toFixed(2),
    pct: d.target > 0 ? Math.round(d.achieved / d.target * 100) : 0,
    finTarget: d.finTarget,
    finAchieved: d.finAchieved,
    finPct: d.finTarget > 0 ? Math.round(d.finAchieved / d.finTarget * 100) : 0
  })).sort((a, b) => b.achieved - a.achieved);

  const tbody = $('kam-summary-body');
  if (tbody) {
    tbody.innerHTML = kamList.map(k => `<tr>
      <td style="font-weight:600">${k.person}</td>
      <td style="font-weight:600">${fT(k.target)}</td>
      <td style="font-weight:600;color:${C.orange}">${fT(k.achieved)}</td>
      <td style="font-weight:700;color:${k.pct >= 100 ? '#10b981' : C.orange}">${k.pct}%</td>
      <td>${fR(k.finTarget)}</td>
      <td>${fR(k.finAchieved)}</td>
      <td style="font-weight:700;color:${k.finPct >= 100 ? '#10b981' : C.orange}">${k.finPct}%</td>
    </tr>`).join('');
  }
}

// ═══════════════════════════════════════════════════════════════
// PANEL 3: DAILY AVERAGE & PERFORMANCE
// ═══════════════════════════════════════════════════════════════
const dtb = $('daily-table-body');
if (dtb) {
  const sortedByAvg = [...clients].sort((a, b) => b.avgDay - a.avgDay);
  dtb.innerHTML = sortedByAvg.map(c => `<tr>
    <td style="font-weight:600">${c.name}</td>
    <td style="color:${C.text3};font-size:.78rem">${c.person}</td>
    <td>${fT(c.target)}</td>
    <td style="font-weight:600">${fT(c.achieved)}</td>
    <td style="font-weight:700;color:${c.pct >= 100 ? '#10b981' : (c.pct >= 70 ? C.orange : C.text3)}">${c.pct}%</td>
    <td style="color:${C.text3}">20</td>
    <td style="font-weight:600;color:${C.orange}">${c.avgDay} T/day</td>
    <td>${fT(c.remaining)}</td>
    <td><span style="display:inline-block;padding:2px 8px;border-radius:4px;font-size:.72rem;font-weight:600;${c.pct >= 100 ? 'background:rgba(16,185,129,0.15);color:#10b981' : (c.pct >= 70 ? 'background:rgba(245,130,32,0.15);color:#f58220' : 'background:#262626;color:#8b8b8b')}">${c.status}</span></td>
  </tr>`).join('');
}

// Avg T/Day Chart
const topAvgClients = [...clients].sort((a, b) => b.avgDay - a.avgDay).slice(0, 10);
safeChart($('avgDayChart'), {
  type: 'bar',
  data: {
    labels: topAvgClients.map(c => c.name),
    datasets: [{
      label: 'Avg T/Day',
      data: topAvgClients.map(c => c.avgDay),
      backgroundColor: C.orange,
      borderRadius: 4
    }]
  },
  options: {
    indexAxis: 'y',
    responsive: true,
    maintainAspectRatio: false,
    scales: {
      x: {
        ticks: {
          callback: v => v + ' T/day'
        }
      }
    },
    plugins: {
      legend: { display: false },
      tooltip: {
        callbacks: {
          label: ctx => ` ${ctx.raw} T/day`
        }
      }
    }
  }
});

// Yesterday Daily Tonnage by Customer (from Daily Tonnage Sheet)
safeChart($('daysChart'), {
  type: 'bar',
  data: {
    labels: dailyTonnage.map(d => d.name),
    datasets: [{
      label: 'Yesterday Tonnage (T)',
      data: dailyTonnage.map(d => d.tonnes),
      backgroundColor: orangeA(0.8),
      borderRadius: 4
    }]
  },
  options: {
    indexAxis: 'y',
    responsive: true,
    maintainAspectRatio: false,
    scales: {
      x: {
        ticks: {
          callback: v => v + ' T'
        }
      }
    },
    plugins: {
      legend: { display: false },
      tooltip: {
        callbacks: {
          label: ctx => ` ${ctx.raw} T (${(ctx.raw / dailyTotalTonnage * 100).toFixed(1)}% of daily total)`
        }
      }
    }
  }
});

// Initialize Tonnage Panel
renderTonnageBars();
renderTonnageCharts();
renderKamSummary();

// Animate Counters
animateCounters();
