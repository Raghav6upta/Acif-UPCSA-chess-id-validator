import { Hono } from 'hono'
import * as xlsx from 'xlsx'
import * as cheerio from 'cheerio'

const app = new Hono()

// Use the bundled index.html (Cloudflare Workers text module or inline)
// For simplicity and since index.html is small, we'll import it as text in wrangler.jsonc or read it
import indexHtml from './index.html'

const AICF_URL = "https://admin.aicf.in/api/players"
const UPCSA_URL = "https://www.upchess.org/view_player.php"

function clean(value: any): string {
  if (value === null || value === undefined) return ""
  return String(value).replace(/\s+/g, " ").trim()
}

function textAfterLabel($: cheerio.CheerioAPI, label: string): string {
  let result = ""
  $("td, th").each((i, el) => {
    const text = clean($(el).text()).toLowerCase()
    if (text === label.toLowerCase()) {
      const nextCell = $(el).next("td, th")
      if (nextCell.length) {
        result = clean(nextCell.text())
      }
    }
  })
  return result
}

async function aicfLookup(playerId: string) {
  const url = `${AICF_URL}?name=${encodeURIComponent(playerId)}&state=0&city=0`
  const res = await fetch(url)
  if (!res.ok) throw new Error(`HTTP ${res.status} from ${url}`)
  
  const payload: any = await res.json()
  const player = payload.data?.find((x: any) => clean(x.aicf_id) === playerId)

  if (!player) {
    return {
      verification_status: "Not Found", aicf_id: "", name: "", gender: "", 
      city: "", district: "", state: "", membership_status: "", membership_expiry: ""
    }
  }

  let expiry = clean(player.membership_expire_at)
  if (expiry) {
    try {
      const d = new Date(expiry.replace("Z", "+00:00"))
      // format to dd-MMM-yyyy (e.g. 05-Jan-2025)
      const months = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]
      expiry = `${String(d.getDate()).padStart(2, '0')}-${months[d.getMonth()]}-${d.getFullYear()}`
    } catch (e) {
      // Keep original if parsing fails
    }
  }

  const nameParts = [clean(player.first_name), clean(player.middle_name), clean(player.last_name)].filter(Boolean)
  const name = nameParts.join(" ")

  return {
    verification_status: "Found",
    aicf_id: clean(player.aicf_id),
    name,
    gender: clean(player.gender),
    city: clean(player.city_name),
    district: clean(player.district_name),
    state: clean(player.state_name),
    membership_status: player.membership_status ? "Active" : "Not Active",
    membership_expiry: expiry
  }
}

async function upcsaLookup(playerId: string) {
  const headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36",
    "Referer": "https://www.upchess.org/",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.9"
  }
  
  const url = `${UPCSA_URL}?id=${encodeURIComponent(playerId)}`
  
  // Try up to 3 times (simplified retry logic for Cloudflare Worker context)
  let pageHtml = ""
  let lastError: any = null
  
  for (let attempt = 0; attempt < 3; attempt++) {
    try {
      const res = await fetch(url, { headers })
      if (!res.ok) throw new Error(`HTTP ${res.status}`)
      pageHtml = await res.text()
      break
    } catch (err) {
      lastError = err
      if (attempt < 2) {
        // Sleep function for TS
        await new Promise(r => setTimeout(r, attempt === 0 ? 8000 : 15000))
      }
    }
  }

  if (!pageHtml) {
    throw new Error(`UPCSA request failed for ${playerId} after 3 attempts: ${lastError?.message}`)
  }

  const $ = cheerio.load(pageHtml)
  const pageText = clean($("body").text())

  const data: any = {
    up_player_id: playerId,
    name: textAfterLabel($, "Name"),
    email: textAfterLabel($, "Email"),
    mobile: textAfterLabel($, "Mobile No"),
    father_name: textAfterLabel($, "Father's Name"),
    gender: textAfterLabel($, "Gender"),
    year_of_birth: textAfterLabel($, "Year of Birth"),
    address: textAfterLabel($, "Address"),
    aicf_id: textAfterLabel($, "AICF ID"),
    fide_id: textAfterLabel($, "FIDE ID"),
    registered_as: textAfterLabel($, "Registered As:"),
    district: "",
    state: "",
    membership_status: "Unknown",
    membership_expiry: ""
  }

  $("td, th").each((i, el) => {
    const text = clean($(el).text())
    if (text.includes("District:") && text.includes("State:")) {
      const m = text.match(/District:\s*(.*?)\s+State:\s*(.*)$/i)
      if (m) {
        data.district = clean(m[1])
        data.state = clean(m[2])
      }
    }
  })

  const inactive = /Not\s*Active!\s*Renew\s*Your\s*Membership/i.test(pageText)
  const activeMatch = pageText.match(/Valid\s*Up\s*To\s*:?\s*([0-9]{1,2}\s+[A-Za-z]+\s+[0-9]{4})/i)

  if (inactive) {
    data.membership_status = "Not Active"
  } else if (activeMatch) {
    data.membership_status = "Active"
    data.membership_expiry = clean(activeMatch[1])
  }

  data.verification_status = data.name ? "Found" : "Not Found"
  return data
}

async function processExcel(fileBuffer: ArrayBuffer, filename: string, mode: string) {
  // Read workbook
  const wb = xlsx.read(fileBuffer, { type: "buffer" })
  const wsName = wb.SheetNames[0]
  const ws = wb.Sheets[wsName]
  
  // Convert sheet to JSON array of arrays (AOA) to manipulate it easily
  const data: any[][] = xlsx.utils.sheet_to_json(ws, { header: 1 })
  if (data.length === 0) throw new Error("Empty Excel file")

  const headers = data[0].map(h => clean(h).toLowerCase())
  const idIdx = headers.findIndex(h => ["playersid", "player id", "playerid", "aicf id", "up player id"].includes(h))
  
  if (idIdx === -1) {
    throw new Error("Could not find a PlayersID / Player ID column in row 1.")
  }

  let newColumns: string[] = []
  if (mode === "aicf") {
    newColumns = ["AICF Verification Status","AICF ID","AICF Name","AICF Gender","AICF City","AICF District","AICF State","AICF Membership Status","AICF Membership Expiry"]
  } else {
    newColumns = ["UPCSA Verification Status","UPCSA Name","UPCSA Father's Name","UPCSA Gender","UPCSA Year of Birth","UPCSA District","UPCSA State","UPCSA Registered As","UPCSA AICF ID","UPCSA FIDE ID","UPCSA Membership Status","UPCSA Membership Expiry","UPCSA Email (masked)","UPCSA Mobile (masked)","UPCSA Address (masked)"]
  }

  // Append new headers
  data[0].push(...newColumns)

  const cache = new Map<string, any>()

  for (let row = 1; row < data.length; row++) {
    // Fill empty cells if the row is shorter than the header length initially
    while (data[row].length <= idIdx) data[row].push("")
    
    const playerId = clean(data[row][idIdx])
    if (!playerId) {
      data[row].push(...Array(newColumns.length).fill(""))
      continue
    }

    if (!cache.has(playerId)) {
      try {
        if (mode === "aicf") {
          cache.set(playerId, await aicfLookup(playerId))
        } else {
          cache.set(playerId, await upcsaLookup(playerId))
        }
      } catch (err: any) {
        cache.set(playerId, { verification_status: `Error: ${err.message || 'Unknown'}` })
      }
    }

    const d = cache.get(playerId)
    let values: string[] = []

    if (mode === "aicf") {
      values = [d.verification_status || "", d.aicf_id || "", d.name || "", d.gender || "", d.city || "", d.district || "", d.state || "", d.membership_status || "", d.membership_expiry || ""]
    } else {
      values = [d.verification_status || "", d.name || "", d.father_name || "", d.gender || "", d.year_of_birth || "", d.district || "", d.state || "", d.registered_as || "", d.aicf_id || "", d.fide_id || "", d.membership_status || "", d.membership_expiry || "", d.email || "", d.mobile || "", d.address || ""]
    }

    data[row].push(...values)
  }

  // Convert back to sheet
  const newWs = xlsx.utils.aoa_to_sheet(data)
  wb.Sheets[wsName] = newWs

  const outBuffer = xlsx.write(wb, { type: "buffer", bookType: "xlsx" })
  const outName = filename.replace(/\.[^/.]+$/, "") + "_verified.xlsx"

  return { buffer: outBuffer, name: outName }
}

app.get("/", (c) => {
  return c.html(indexHtml)
})

app.post("/process", async (c) => {
  const body = await c.req.parseBody()
  const file = body['file'] as File
  const mode = (body['mode'] as string || "").toLowerCase()

  if (!file) {
    return c.json({ error: "No Excel file uploaded." }, 400)
  }
  if (mode !== "aicf" && mode !== "upcsa") {
    return c.json({ error: "Invalid mode." }, 400)
  }

  try {
    const fileBuffer = await file.arrayBuffer()
    const { buffer, name } = await processExcel(fileBuffer, file.name, mode)
    
    return new Response(buffer, {
      headers: {
        'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'Content-Disposition': `attachment; filename="${name}"`
      }
    })
  } catch (err: any) {
    return c.json({ error: `${err.name || 'Error'}: ${err.message}` }, 500)
  }
})

export default app
