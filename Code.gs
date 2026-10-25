// ============================================================
// COMPETITOR BLOG & RESOURCE TRACKER v2
// Scrapes competitor blogs and resource pages monthly,
// summarizes new content via AI, reports to Google Sheets + Slack.
// Configure for your organization using Script Properties and the Competitors sheet.
// ============================================================

var CONFIG = {
  // --- PRIVATE SCRIPT PROPERTIES (placeholders until configured) ---
  groqApiKey: PropertiesService.getScriptProperties().getProperty('GROQ_API_KEY') || 'YOUR_GROQ_API_KEY',
  groqApiKey2: PropertiesService.getScriptProperties().getProperty('GROQ_API_KEY_SECONDARY') || 'YOUR_OPTIONAL_SECONDARY_GROQ_API_KEY',

  // --- SLACK ---
  slackWebhookUrl: PropertiesService.getScriptProperties().getProperty('SLACK_WEBHOOK_URL') || 'YOUR_SLACK_WEBHOOK_URL',

  // --- SHEET NAMES ---
  competitorsSheet: 'Competitors',
  reportSheet: 'Monthly report',
  archiveSheet: 'Archive',

  // --- SETTINGS ---
  maxPostsPerCompetitor: 20,       // max items to extract per source URL
  maxPostsToSummarize: 30,         // total posts to send for AI summary (to stay within limits)
  requestDelay: 500,               // ms between HTTP requests
  aiRetryDelay: 20000,             // ms to wait before retrying a rate-limited AI call
  aiMaxRetries: 3,                 // max retries per AI batch
  summaryBatchSize: 3,             // posts per AI call (smaller = less likely to hit token limits)
};


// ============================================================
// MENU
// ============================================================
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Blog tracker')
    .addItem('Run monthly scan now', 'runBlogTracker')
    .addItem('Set up monthly trigger', 'setupMonthlyTrigger')
    .addItem('Remove triggers', 'removeTriggers')
    .addToMenu();
}


// ============================================================
// TRIGGER MANAGEMENT
// ============================================================
function setupMonthlyTrigger() {
  removeTriggers();
  ScriptApp.newTrigger('runBlogTracker')
    .timeBased()
    .onMonthDay(1)
    .atHour(9)
    .create();

  SpreadsheetApp.getUi().alert(
    'Trigger set',
    'The blog tracker will now run automatically on the 1st of every month during the 9 AM hour in the project time zone.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function removeTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'runBlogTracker') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}


// ============================================================
// MAIN FUNCTION
// ============================================================
function runBlogTracker() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui;
  try { ui = SpreadsheetApp.getUi(); } catch (e) { ui = null; }

  // 1. Read competitors
  var competitors = readCompetitors(ss);
  if (competitors.length === 0) {
    if (ui) ui.alert('Error', 'No competitors found in the "' + CONFIG.competitorsSheet + '" sheet.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  // 2. Determine target month
  var targetDate = getTargetMonth();
  var monthLabel = Utilities.formatDate(targetDate, Session.getScriptTimeZone(), 'MMMM yyyy');
  Logger.log('Scanning for content published in: ' + monthLabel);

  // 3. Scrape all competitor pages
  var allItems = [];
  for (var c = 0; c < competitors.length; c++) {
    var comp = competitors[c];
    Logger.log('--- Scraping: ' + comp.name + ' ---');

    // Scrape blog
    if (comp.blogUrl) {
      var blogItems = scrapePage(comp.blogUrl, comp.name, 'Blog', comp.rssUrl);
      logDateSummary(comp.name, 'Blog', blogItems);
      allItems = allItems.concat(blogItems);
    }

    // Scrape resources
    if (comp.resourceUrl) {
      var resourceItems = scrapePage(comp.resourceUrl, comp.name, 'Resource', null);
      logDateSummary(comp.name, 'Resource', resourceItems);
      allItems = allItems.concat(resourceItems);
    }

    Utilities.sleep(CONFIG.requestDelay);
  }

  Logger.log('=== Total items found: ' + allItems.length + ' ===');

  // 4. For items without dates, try fetching the individual page to extract a date
  var undatedCount = 0;
  for (var ud = 0; ud < allItems.length; ud++) {
    if (!allItems[ud].date && allItems[ud].url) {
      var fetchedDate = fetchDateFromPage(allItems[ud].url);
      if (fetchedDate) {
        allItems[ud].date = fetchedDate.date;
        allItems[ud].dateRaw = fetchedDate.raw;
        allItems[ud].source = allItems[ud].source + '+page';
      } else {
        undatedCount++;
      }
    }
    // Don't burn too much time on this; cap at a reasonable number
    if (ud > 60) break;
  }
  Logger.log('Items still without dates after page fetch: ' + undatedCount);

  // 5. Filter to target month
  //    For dated items: only include if they match the target month
  //    For undated items: keep only the first few per competitor per source type,
  //    since gallery pages list newest content first
  var MAX_UNDATED_PER_COMPETITOR_PER_SOURCE = 5;

  var filteredItems = [];
  var undatedItems = [];
  var undatedCounts = {}; // tracks count per "competitor|sourceType" key

  // Items are already in page order (top of gallery = newest), so we process them
  // in the order they were scraped
  for (var fi = 0; fi < allItems.length; fi++) {
    var item = allItems[fi];
    if (item.date) {
      if (item.date.getMonth() === targetDate.getMonth() && item.date.getFullYear() === targetDate.getFullYear()) {
        filteredItems.push(item);
      }
    } else {
      var countKey = item.competitor + '|' + item.contentType;
      if (!undatedCounts[countKey]) undatedCounts[countKey] = 0;

      if (undatedCounts[countKey] < MAX_UNDATED_PER_COMPETITOR_PER_SOURCE) {
        undatedItems.push(item);
        undatedCounts[countKey]++;
      }
    }
  }

  Logger.log('Dated items matching ' + monthLabel + ': ' + filteredItems.length);
  Logger.log('Undated items (top ' + MAX_UNDATED_PER_COMPETITOR_PER_SOURCE + ' per competitor per source): ' + undatedItems.length);

  // Combine: dated matches first, then undated
  var allReportItems = filteredItems.concat(undatedItems);

  // 6. Generate AI summaries for blog posts (skip gated resources)
  var itemsToSummarize = allReportItems.filter(function(item) {
    return item.contentType === 'Blog' && item.url;
  }).slice(0, CONFIG.maxPostsToSummarize);

  if (itemsToSummarize.length > 0) {
    Logger.log('Generating AI summaries for ' + itemsToSummarize.length + ' blog posts...');
    var summaries = generateSummaries(itemsToSummarize);
    for (var s = 0; s < summaries.length; s++) {
      for (var f = 0; f < allReportItems.length; f++) {
        if (allReportItems[f].url === summaries[s].url) {
          allReportItems[f].summary = summaries[s].summary;
          break;
        }
      }
    }
  }

  // 7. For resource items without summaries, try meta description
  for (var ri = 0; ri < allReportItems.length; ri++) {
    if (allReportItems[ri].contentType !== 'Blog' && !allReportItems[ri].summary && allReportItems[ri].url) {
      var metaDesc = fetchMetaDescription(allReportItems[ri].url);
      if (metaDesc) {
        allReportItems[ri].summary = metaDesc;
      }
      Utilities.sleep(300);
    }
  }

  // 8. Classify content types for resource items
  classifyContentTypes(allReportItems);

  // 9. Write to Monthly report sheet
  writeMonthlyReport(ss, allReportItems, monthLabel);

  // 10. Append to Archive sheet
  appendToArchive(ss, allReportItems, monthLabel);

  // 11. Send Slack notification
  sendSlackReport(allReportItems, monthLabel, ss.getUrl());

  // 12. Done
  Logger.log('Blog tracker complete. Found ' + allReportItems.length + ' items (' + filteredItems.length + ' dated, ' + undatedItems.length + ' undated) for ' + monthLabel);
  if (ui) {
    ui.alert(
      'Scan complete',
      'Found ' + allReportItems.length + ' content items for ' + monthLabel + '.\n' +
      '(' + filteredItems.length + ' with confirmed dates, ' + undatedItems.length + ' undated, top ' + MAX_UNDATED_PER_COMPETITOR_PER_SOURCE + ' per competitor)\n\n' +
      'Results are in the "' + CONFIG.reportSheet + '" tab.\nSlack notification sent.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}


// ============================================================
// LOGGING HELPER
// ============================================================
function logDateSummary(compName, sourceType, items) {
  var dated = 0;
  var undated = 0;
  for (var i = 0; i < items.length; i++) {
    if (items[i].date) { dated++; } else { undated++; }
  }
  Logger.log(compName + ' (' + sourceType + '): ' + items.length + ' items found (' + dated + ' with dates, ' + undated + ' undated)');
  // Log first 3 items for debugging
  for (var j = 0; j < Math.min(3, items.length); j++) {
    Logger.log('  Sample: "' + items[j].title.substring(0, 60) + '" | date: ' + (items[j].date || 'NONE') + ' | source: ' + items[j].source);
  }
}


// ============================================================
// READ COMPETITORS FROM SHEET
// ============================================================
function readCompetitors(ss) {
  var sheet = ss.getSheetByName(CONFIG.competitorsSheet);
  if (!sheet) return [];

  var data = sheet.getDataRange().getValues();
  var competitors = [];

  for (var i = 1; i < data.length; i++) {
    var name = data[i][0] ? data[i][0].toString().trim() : '';
    var blogUrl = data[i][1] ? data[i][1].toString().trim() : '';
    var resourceUrl = data[i][2] ? data[i][2].toString().trim() : '';
    var rssUrl = data[i][3] ? data[i][3].toString().trim() : '';

    if (name && (blogUrl || resourceUrl)) {
      competitors.push({
        name: name,
        blogUrl: blogUrl,
        resourceUrl: resourceUrl,
        rssUrl: rssUrl,
      });
    }
  }

  return competitors;
}


// ============================================================
// TARGET MONTH
// ============================================================
function getTargetMonth() {
  var now = new Date();
  if (now.getDate() <= 3) {
    return new Date(now.getFullYear(), now.getMonth() - 1, 1);
  } else {
    return new Date(now.getFullYear(), now.getMonth(), 1);
  }
}


// ============================================================
// FETCH DATE FROM INDIVIDUAL POST PAGE
// (used when gallery page didn't provide a date)
// ============================================================
function fetchDateFromPage(url) {
  try {
    var html = fetchPage(url);
    if (!html) return null;

    // 1. JSON-LD datePublished
    var jsonLdPattern = /<script[^>]*type\s*=\s*["']application\/ld\+json["'][^>]*>([\s\S]*?)<\/script>/gi;
    var jMatch;
    while ((jMatch = jsonLdPattern.exec(html)) !== null) {
      try {
        var jData = JSON.parse(jMatch[1]);
        var dataList = Array.isArray(jData) ? jData : [jData];
        for (var d = 0; d < dataList.length; d++) {
          var obj = dataList[d];
          if (obj.datePublished) {
            var pd = parseDate(obj.datePublished);
            if (pd) return { date: pd, raw: obj.datePublished };
          }
          if (obj['@graph']) {
            for (var g = 0; g < obj['@graph'].length; g++) {
              if (obj['@graph'][g].datePublished) {
                var gd = parseDate(obj['@graph'][g].datePublished);
                if (gd) return { date: gd, raw: obj['@graph'][g].datePublished };
              }
            }
          }
        }
      } catch (e) {}
    }

    // 2. <time datetime="...">
    var timeMatch = html.match(/<time[^>]*datetime=["']([^"']+)["']/i);
    if (timeMatch) {
      var td = parseDate(timeMatch[1]);
      if (td) return { date: td, raw: timeMatch[1] };
    }

    // 3. <meta property="article:published_time">
    var metaMatch = html.match(/<meta[^>]*property=["']article:published_time["'][^>]*content=["']([^"']+)["']/i);
    if (metaMatch) {
      var md = parseDate(metaMatch[1]);
      if (md) return { date: md, raw: metaMatch[1] };
    }

    // 4. Visible date patterns near the top of the page (first 5000 chars of body)
    var bodyMatch = html.match(/<body[\s\S]*$/i);
    var bodyTop = bodyMatch ? bodyMatch[0].substring(0, 5000) : html.substring(0, 5000);
    var visibleText = stripHTML(bodyTop);

    // "July 6, 2026" or "Jul 6, 2026"
    var textDateMatch = visibleText.match(/(?:January|February|March|April|May|June|July|August|September|October|November|December|Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\.?\s+\d{1,2},?\s+\d{4}/i);
    if (textDateMatch) {
      var vd = parseDate(textDateMatch[0]);
      if (vd) return { date: vd, raw: textDateMatch[0] };
    }

    // "6 July 2026"
    var textDateMatch2 = visibleText.match(/\d{1,2}\s+(?:January|February|March|April|May|June|July|August|September|October|November|December|Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\.?\s+\d{4}/i);
    if (textDateMatch2) {
      var vd2 = parseDate(textDateMatch2[0]);
      if (vd2) return { date: vd2, raw: textDateMatch2[0] };
    }

    return null;
  } catch (e) {
    return null;
  }
}


// ============================================================
// SCRAPE A PAGE
// ============================================================
function scrapePage(pageUrl, competitorName, sourceType, rssUrl) {
  var items = [];

  // Try RSS first
  if (rssUrl) {
    var rssItems = parseRSS(rssUrl, competitorName, sourceType);
    if (rssItems.length > 0) {
      Logger.log('Got ' + rssItems.length + ' items from RSS for ' + competitorName);
      return rssItems;
    }
  }

  // Auto-detect RSS
  if (sourceType === 'Blog') {
    var detectedRss = tryDetectRSS(pageUrl);
    if (detectedRss) {
      var rssItems2 = parseRSS(detectedRss, competitorName, sourceType);
      if (rssItems2.length > 0) {
        Logger.log('Auto-detected RSS for ' + competitorName + ': ' + detectedRss);
        return rssItems2;
      }
    }
  }

  // HTML scraping
  var html = fetchPage(pageUrl);
  if (!html) return items;

  // Try JSON-LD first
  var jsonLdItems = extractFromJsonLd(html, competitorName, sourceType, pageUrl);
  if (jsonLdItems.length > 0) {
    Logger.log('Got ' + jsonLdItems.length + ' items from JSON-LD for ' + competitorName);
    return jsonLdItems.slice(0, CONFIG.maxPostsPerCompetitor);
  }

  // Fall back to HTML
  var htmlItems = extractFromHTML(html, competitorName, sourceType, pageUrl);
  Logger.log('Got ' + htmlItems.length + ' items from HTML for ' + competitorName);
  return htmlItems.slice(0, CONFIG.maxPostsPerCompetitor);
}


// ============================================================
// RSS
// ============================================================
function tryDetectRSS(pageUrl) {
  var domain = pageUrl.replace(/https?:\/\//, '').split('/')[0];
  var base = 'https://' + domain;
  var feedPaths = ['/feed', '/rss', '/blog/feed', '/blog/rss', '/feed.xml', '/rss.xml', '/atom.xml', '/blog/feed.xml'];

  for (var i = 0; i < feedPaths.length; i++) {
    try {
      var response = UrlFetchApp.fetch(base + feedPaths[i], {
        muteHttpExceptions: true,
        followRedirects: true,
        headers: { 'User-Agent': 'Mozilla/5.0 (compatible; BlogTracker/1.0)' },
      });
      if (response.getResponseCode() === 200) {
        var content = response.getContentText().trim();
        if (content.indexOf('<rss') !== -1 || content.indexOf('<feed') !== -1 || content.indexOf('<channel') !== -1) {
          return base + feedPaths[i];
        }
      }
    } catch (e) {}
    Utilities.sleep(200);
  }
  return null;
}

function parseRSS(rssUrl, competitorName, sourceType) {
  var items = [];
  try {
    var response = UrlFetchApp.fetch(rssUrl, {
      muteHttpExceptions: true,
      followRedirects: true,
      headers: { 'User-Agent': 'Mozilla/5.0 (compatible; BlogTracker/1.0)' },
    });
    if (response.getResponseCode() !== 200) return items;

    var content = response.getContentText();
    var doc = XmlService.parse(content);
    var root = doc.getRootElement();

    // RSS 2.0
    var ns = root.getNamespace();
    var channel = root.getChild('channel', ns) || root.getChild('channel');

    if (channel) {
      var rssItems = channel.getChildren('item', ns) || channel.getChildren('item');
      for (var i = 0; i < rssItems.length && i < CONFIG.maxPostsPerCompetitor; i++) {
        var item = rssItems[i];
        var title = getChildText(item, 'title', ns);
        var link = getChildText(item, 'link', ns);
        var pubDate = getChildText(item, 'pubDate', ns);
        var description = getChildText(item, 'description', ns);

        if (title) {
          items.push({
            competitor: competitorName,
            title: cleanText(title),
            url: link || '',
            date: pubDate ? parseDate(pubDate) : null,
            dateRaw: pubDate || '',
            contentType: sourceType,
            summary: description ? cleanText(stripHTML(description)).substring(0, 300) : '',
            source: 'RSS',
          });
        }
      }
    }

    // Atom
    if (items.length === 0) {
      var atomNs = XmlService.getNamespace('http://www.w3.org/2005/Atom');
      var entries = root.getChildren('entry', atomNs) || root.getChildren('entry');
      for (var j = 0; j < entries.length && j < CONFIG.maxPostsPerCompetitor; j++) {
        var entry = entries[j];
        var aTitle = getChildText(entry, 'title', atomNs);
        var aLink = '';
        var linkEl = entry.getChild('link', atomNs) || entry.getChild('link');
        if (linkEl) aLink = linkEl.getAttribute('href') ? linkEl.getAttribute('href').getValue() : '';
        var aDate = getChildText(entry, 'published', atomNs) || getChildText(entry, 'updated', atomNs);
        var aSummary = getChildText(entry, 'summary', atomNs) || getChildText(entry, 'content', atomNs);

        if (aTitle) {
          items.push({
            competitor: competitorName,
            title: cleanText(aTitle),
            url: aLink,
            date: aDate ? parseDate(aDate) : null,
            dateRaw: aDate || '',
            contentType: sourceType,
            summary: aSummary ? cleanText(stripHTML(aSummary)).substring(0, 300) : '',
            source: 'RSS',
          });
        }
      }
    }
  } catch (e) {
    Logger.log('RSS parse error for ' + rssUrl + ': ' + e.message);
  }
  return items;
}


// ============================================================
// JSON-LD EXTRACTION
// ============================================================
function extractFromJsonLd(html, competitorName, sourceType, pageUrl) {
  var items = [];
  var scriptPattern = /<script[^>]*type\s*=\s*["']application\/ld\+json["'][^>]*>([\s\S]*?)<\/script>/gi;
  var match;

  while ((match = scriptPattern.exec(html)) !== null) {
    try {
      var data = JSON.parse(match[1].trim());
      var dataList = Array.isArray(data) ? data : [data];

      for (var d = 0; d < dataList.length; d++) {
        var obj = dataList[d];

        if (obj['@type'] === 'BlogPosting' || obj['@type'] === 'Article' || obj['@type'] === 'NewsArticle') {
          items.push({
            competitor: competitorName,
            title: cleanText(obj.headline || obj.name || ''),
            url: obj.url || obj.mainEntityOfPage || '',
            date: obj.datePublished ? parseDate(obj.datePublished) : null,
            dateRaw: obj.datePublished || '',
            contentType: sourceType,
            summary: obj.description ? cleanText(obj.description).substring(0, 300) : '',
            source: 'JSON-LD',
          });
        }

        if (obj['@type'] === 'ItemList' && obj.itemListElement) {
          for (var il = 0; il < obj.itemListElement.length; il++) {
            var inner = obj.itemListElement[il].item || obj.itemListElement[il];
            if (inner.name || inner.headline) {
              items.push({
                competitor: competitorName,
                title: cleanText(inner.headline || inner.name || ''),
                url: inner.url || '',
                date: inner.datePublished ? parseDate(inner.datePublished) : null,
                dateRaw: inner.datePublished || '',
                contentType: sourceType,
                summary: inner.description ? cleanText(inner.description).substring(0, 300) : '',
                source: 'JSON-LD',
              });
            }
          }
        }

        if (obj['@graph']) {
          for (var g = 0; g < obj['@graph'].length; g++) {
            var gObj = obj['@graph'][g];
            if (gObj['@type'] === 'BlogPosting' || gObj['@type'] === 'Article') {
              items.push({
                competitor: competitorName,
                title: cleanText(gObj.headline || gObj.name || ''),
                url: gObj.url || gObj.mainEntityOfPage || '',
                date: gObj.datePublished ? parseDate(gObj.datePublished) : null,
                dateRaw: gObj.datePublished || '',
                contentType: sourceType,
                summary: gObj.description ? cleanText(gObj.description).substring(0, 300) : '',
                source: 'JSON-LD',
              });
            }
          }
        }
      }
    } catch (e) {}
  }
  return items;
}


// ============================================================
// HTML EXTRACTION (fallback)
// ============================================================
function extractFromHTML(html, competitorName, sourceType, pageUrl) {
  var items = [];
  var domain = pageUrl.replace(/https?:\/\//, '').split('/')[0];
  var base = 'https://' + domain;

  // Strategy 1: <article> blocks
  var articlePattern = /<article[^>]*>([\s\S]*?)<\/article>/gi;
  var articleMatch;
  while ((articleMatch = articlePattern.exec(html)) !== null && items.length < CONFIG.maxPostsPerCompetitor) {
    var item = extractItemFromBlock(articleMatch[1], competitorName, sourceType, base);
    if (item && item.title) items.push(item);
  }

  // Strategy 2: Common blog card class patterns
  if (items.length === 0) {
    var cardPatterns = [
      /<div[^>]*class="[^"]*(?:blog-?(?:post|card|item|entry)|post-?(?:card|item|entry)|article-?(?:card|item)|resource-?(?:card|item))[^"]*"[^>]*>([\s\S]*?)<\/div>\s*(?:<\/div>)?/gi,
      /<a[^>]*class="[^"]*(?:blog-?(?:post|card|item|entry)|post-?(?:card|item|entry)|article-?(?:card|item)|resource-?(?:card|item))[^"]*"[^>]*>([\s\S]*?)<\/a>/gi,
      /<li[^>]*class="[^"]*(?:blog|post|article|resource)[^"]*"[^>]*>([\s\S]*?)<\/li>/gi,
    ];

    for (var cp = 0; cp < cardPatterns.length && items.length === 0; cp++) {
      var cardMatch;
      while ((cardMatch = cardPatterns[cp].exec(html)) !== null && items.length < CONFIG.maxPostsPerCompetitor) {
        var cardItem = extractItemFromBlock(cardMatch[0], competitorName, sourceType, base);
        if (cardItem && cardItem.title) items.push(cardItem);
      }
    }
  }

  // Strategy 3: All links that look like blog or resource posts
  if (items.length === 0) {
    var linkPattern = /<a[^>]*href=["']([^"']*(?:\/blog\/|\/post\/|\/articles\/|\/resources\/|\/guides\/|\/learning-center\/|\/reports\/|\/webflow-guides\/)[^"']+)["'][^>]*>([\s\S]*?)<\/a>/gi;
    var linkMatch;
    var seenUrls = {};
    while ((linkMatch = linkPattern.exec(html)) !== null && items.length < CONFIG.maxPostsPerCompetitor) {
      var href = linkMatch[1];
      var linkText = cleanText(stripHTML(linkMatch[2]));

      if (href.indexOf('http') !== 0) {
        href = href.indexOf('/') === 0 ? base + href : base + '/' + href;
      }

      if (seenUrls[href] || href.indexOf('#') !== -1 || linkText.length < 5 || linkText.length > 300) continue;
      seenUrls[href] = true;

      items.push({
        competitor: competitorName,
        title: linkText,
        url: href,
        date: null,
        dateRaw: '',
        contentType: sourceType,
        summary: '',
        source: 'HTML-link',
      });
    }
  }

  // Strategy 4: As a broader fallback, find any internal links with substantial anchor text
  // that include path segments suggesting content (not nav links)
  if (items.length === 0) {
    var broadPattern = /<a[^>]*href=["']((?:https?:\/\/[^"']*)?\/[^"']+)["'][^>]*>([\s\S]*?)<\/a>/gi;
    var broadMatch;
    var seenBroad = {};
    while ((broadMatch = broadPattern.exec(html)) !== null && items.length < CONFIG.maxPostsPerCompetitor) {
      var bHref = broadMatch[1];
      var bText = cleanText(stripHTML(broadMatch[2]));

      if (bHref.indexOf('http') !== 0) {
        bHref = bHref.indexOf('/') === 0 ? base + bHref : base + '/' + bHref;
      }

      // Must be on the same domain
      if (bHref.indexOf(domain) === -1) continue;
      // Must have substantial text (a real title, not "Read more")
      if (bText.length < 15 || bText.length > 300) continue;
      // Skip common nav/footer links
      if (/^(home|about|contact|pricing|login|sign|blog|resources|services|careers|faq|privacy|terms)/i.test(bText)) continue;
      if (seenBroad[bHref]) continue;
      seenBroad[bHref] = true;

      items.push({
        competitor: competitorName,
        title: bText,
        url: bHref,
        date: null,
        dateRaw: '',
        contentType: sourceType,
        summary: '',
        source: 'HTML-broad',
      });
    }
  }

  // Try to extract dates for items from URL patterns
  for (var di = 0; di < items.length; di++) {
    if (!items[di].date && items[di].url) {
      var dateFromUrl = extractDateFromUrl(items[di].url);
      if (dateFromUrl) {
        items[di].date = dateFromUrl;
        items[di].dateRaw = 'from URL';
      }
    }
  }

  return items;
}

function extractItemFromBlock(block, competitorName, sourceType, baseUrl) {
  // Title
  var title = '';
  var titlePatterns = [
    /<h[1234][^>]*>([\s\S]*?)<\/h[1234]>/i,
    /<a[^>]*>([\s\S]*?)<\/a>/i,
    /<strong[^>]*>([\s\S]*?)<\/strong>/i,
  ];
  for (var t = 0; t < titlePatterns.length; t++) {
    var tMatch = block.match(titlePatterns[t]);
    if (tMatch) {
      title = cleanText(stripHTML(tMatch[1]));
      if (title.length >= 5 && title.length <= 300) break;
      title = '';
    }
  }

  // URL
  var url = '';
  var urlMatch = block.match(/<a[^>]*href=["']([^"'#]+)["']/i);
  if (urlMatch) {
    url = urlMatch[1];
    if (url.indexOf('http') !== 0) {
      url = url.indexOf('/') === 0 ? baseUrl + url : baseUrl + '/' + url;
    }
  }

  // Date: <time datetime>
  var date = null;
  var dateRaw = '';
  var timeMatch = block.match(/<time[^>]*datetime=["']([^"']+)["']/i);
  if (timeMatch) {
    date = parseDate(timeMatch[1]);
    dateRaw = timeMatch[1];
  }

  // Date: visible text "Month Day, Year"
  if (!date) {
    var dateTextMatch = block.match(/(?:January|February|March|April|May|June|July|August|September|October|November|December|Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\.?\s+\d{1,2},?\s+\d{4}/i);
    if (dateTextMatch) {
      date = parseDate(dateTextMatch[0]);
      dateRaw = dateTextMatch[0];
    }
  }

  // Date: "Day Month Year" format
  if (!date) {
    var dateTextMatch2 = block.match(/\d{1,2}\s+(?:January|February|March|April|May|June|July|August|September|October|November|December|Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\.?\s+\d{4}/i);
    if (dateTextMatch2) {
      date = parseDate(dateTextMatch2[0]);
      dateRaw = dateTextMatch2[0];
    }
  }

  // Date: numeric "2026-07-06"
  if (!date) {
    var numDateMatch = block.match(/\b(\d{4}[-\/]\d{1,2}[-\/]\d{1,2})\b/);
    if (numDateMatch) {
      date = parseDate(numDateMatch[1]);
      dateRaw = numDateMatch[1];
    }
  }

  return {
    competitor: competitorName,
    title: title,
    url: url,
    date: date,
    dateRaw: dateRaw,
    contentType: sourceType,
    summary: '',
    source: 'HTML',
  };
}


// ============================================================
// DATE UTILITIES
// ============================================================
function parseDate(dateStr) {
  if (!dateStr) return null;
  try {
    var d = new Date(dateStr);
    if (!isNaN(d.getTime()) && d.getFullYear() > 2020) return d;
  } catch (e) {}

  var patterns = [
    /(\w+)\s+(\d{1,2}),?\s+(\d{4})/,
    /(\d{1,2})\s+(\w+)\s+(\d{4})/,
    /(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})/,
    /(\d{1,2})[-\/](\d{1,2})[-\/](\d{4})/,
  ];

  for (var p = 0; p < patterns.length; p++) {
    var m = dateStr.match(patterns[p]);
    if (m) {
      try {
        var d2 = new Date(m[0]);
        if (!isNaN(d2.getTime()) && d2.getFullYear() > 2020) return d2;
      } catch (e2) {}
    }
  }
  return null;
}

function extractDateFromUrl(url) {
  var urlDateMatch = url.match(/\/(\d{4})\/(\d{1,2})(?:\/(\d{1,2}))?/);
  if (urlDateMatch) {
    var year = parseInt(urlDateMatch[1]);
    var month = parseInt(urlDateMatch[2]) - 1;
    var day = urlDateMatch[3] ? parseInt(urlDateMatch[3]) : 1;
    if (year > 2020 && year < 2030 && month >= 0 && month <= 11) {
      return new Date(year, month, day);
    }
  }
  return null;
}


// ============================================================
// CONTENT TYPE CLASSIFICATION
// ============================================================
function classifyContentTypes(items) {
  var guideKeywords = ['guide', 'ebook', 'e-book', 'whitepaper', 'white paper', 'handbook', 'playbook'];
  var templateKeywords = ['template', 'swipe file', 'checklist', 'cheat sheet', 'worksheet', 'toolkit', 'kit'];
  var reportKeywords = ['report', 'survey', 'benchmark', 'research', 'data', 'statistics', 'state of'];
  var webinarKeywords = ['webinar', 'recording', 'watch', 'on-demand', 'replay'];

  for (var i = 0; i < items.length; i++) {
    if (items[i].contentType === 'Blog') continue;

    var combined = ((items[i].title || '') + ' ' + (items[i].summary || '') + ' ' + (items[i].url || '')).toLowerCase();

    if (matchesAny(combined, guideKeywords)) {
      items[i].contentType = 'Guide';
    } else if (matchesAny(combined, templateKeywords)) {
      items[i].contentType = 'Template';
    } else if (matchesAny(combined, reportKeywords)) {
      items[i].contentType = 'Report';
    } else if (matchesAny(combined, webinarKeywords)) {
      items[i].contentType = 'Webinar';
    } else {
      items[i].contentType = 'Lead magnet';
    }
  }
}

function matchesAny(text, keywords) {
  for (var k = 0; k < keywords.length; k++) {
    if (text.indexOf(keywords[k]) !== -1) return true;
  }
  return false;
}


// ============================================================
// AI SUMMARIES (with retry logic)
// ============================================================
function generateSummaries(items) {
  var results = [];
  var batchContent = [];

  for (var i = 0; i < items.length; i++) {
    var html = fetchPage(items[i].url);
    if (!html) {
      results.push({ url: items[i].url, summary: items[i].summary || '(could not fetch page)' });
      continue;
    }

    var mainText = extractMainContent(html);
    if (mainText.length < 50) {
      results.push({ url: items[i].url, summary: items[i].summary || '(insufficient content)' });
      continue;
    }

    batchContent.push({
      url: items[i].url,
      title: items[i].title,
      text: mainText.substring(0, 1500),
    });

    Utilities.sleep(CONFIG.requestDelay);
  }

  // Send in smaller batches with retry
  for (var b = 0; b < batchContent.length; b += CONFIG.summaryBatchSize) {
    var batch = batchContent.slice(b, b + CONFIG.summaryBatchSize);
    var prompt = buildSummaryPrompt(batch);
    var aiResponse = callAIWithRetry(prompt);

    if (aiResponse && Array.isArray(aiResponse)) {
      for (var a = 0; a < aiResponse.length && a < batch.length; a++) {
        results.push({
          url: aiResponse[a].url || batch[a].url,
          summary: aiResponse[a].summary || '',
        });
      }
    } else {
      for (var fb = 0; fb < batch.length; fb++) {
        results.push({
          url: batch[fb].url,
          summary: batch[fb].text.substring(0, 200) + '...',
        });
      }
    }

    // Small delay between batches to avoid rate limits
    if (b + CONFIG.summaryBatchSize < batchContent.length) {
      Utilities.sleep(3000);
    }
  }

  return results;
}

function buildSummaryPrompt(batch) {
  var prompt = 'You are a content analyst. For each blog post below, write a concise 2-3 sentence outline summary. ';
  prompt += 'Focus on the key takeaway and main topics covered.\n\n';

  for (var i = 0; i < batch.length; i++) {
    prompt += '--- POST ' + (i + 1) + ' ---\n';
    prompt += 'Title: ' + batch[i].title + '\n';
    prompt += 'URL: ' + batch[i].url + '\n';
    prompt += 'Content excerpt:\n' + batch[i].text + '\n\n';
  }

  prompt += '\nRespond ONLY in valid JSON. No markdown, no backticks, no preamble.\n';
  prompt += 'Use this exact structure:\n';
  prompt += '[{"url": "...", "summary": "..."}]\n';
  prompt += 'Return one object per post, in the same order.\n';

  return prompt;
}

function extractMainContent(html) {
  var cleaned = html
    .replace(/<script[\s\S]*?<\/script>/gi, '')
    .replace(/<style[\s\S]*?<\/style>/gi, '')
    .replace(/<nav[\s\S]*?<\/nav>/gi, '')
    .replace(/<header[\s\S]*?<\/header>/gi, '')
    .replace(/<footer[\s\S]*?<\/footer>/gi, '')
    .replace(/<aside[\s\S]*?<\/aside>/gi, '');

  var mainMatch = cleaned.match(/<(?:main|article)[^>]*>([\s\S]*?)<\/(?:main|article)>/i);
  if (mainMatch) return cleanText(stripHTML(mainMatch[1]));

  var contentMatch = cleaned.match(/<div[^>]*class="[^"]*(?:post-?content|article-?(?:content|body)|entry-?content|blog-?(?:content|body))[^"]*"[^>]*>([\s\S]*?)<\/div>/i);
  if (contentMatch) return cleanText(stripHTML(contentMatch[1]));

  var paragraphs = [];
  var pPattern = /<p[^>]*>([\s\S]*?)<\/p>/gi;
  var pMatch;
  while ((pMatch = pPattern.exec(cleaned)) !== null) {
    var pText = cleanText(stripHTML(pMatch[1]));
    if (pText.length > 30) paragraphs.push(pText);
  }
  return paragraphs.join(' ').substring(0, 3000);
}

function fetchMetaDescription(url) {
  try {
    var html = fetchPage(url);
    if (!html) return null;

    var ogMatch = html.match(/<meta[^>]*property=["']og:description["'][^>]*content=["']([^"']+)["']/i);
    if (ogMatch) return cleanText(ogMatch[1]);

    var metaMatch = html.match(/<meta[^>]*name=["']description["'][^>]*content=["']([^"']+)["']/i);
    if (metaMatch) return cleanText(metaMatch[1]);

    return null;
  } catch (e) {
    return null;
  }
}


// ============================================================
// AI API CALLS WITH RETRY
// ============================================================
function callAIWithRetry(prompt) {
  for (var attempt = 0; attempt < CONFIG.aiMaxRetries; attempt++) {
    // Try primary Groq key
    var result1 = callGroqWithKey(prompt, CONFIG.groqApiKey, 'Groq-primary');
    if (result1) return result1;

    // Try secondary Groq key
    var result2 = callGroqWithKey(prompt, CONFIG.groqApiKey2, 'Groq-fallback');
    if (result2) return result2;

    // Both failed, wait and retry
    if (attempt < CONFIG.aiMaxRetries - 1) {
      Logger.log('Both Groq keys failed on attempt ' + (attempt + 1) + '. Waiting ' + (CONFIG.aiRetryDelay / 1000) + 's before retry...');
      Utilities.sleep(CONFIG.aiRetryDelay);
    }
  }

  Logger.log('All AI retries exhausted. Summaries will use text excerpts instead.');
  return null;
}

function callGroqWithKey(prompt, apiKey, label) {
  if (!apiKey || apiKey.indexOf('YOUR_') === 0) return null;
  try {
    var url = 'https://api.groq.com/openai/v1/chat/completions';
    var payload = {
      model: 'llama-3.3-70b-versatile',
      messages: [{ role: 'user', content: prompt }],
      temperature: 0.3,
      max_tokens: 2000,
    };
    var options = {
      method: 'post',
      contentType: 'application/json',
      headers: { 'Authorization': 'Bearer ' + apiKey },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true,
    };

    var response = UrlFetchApp.fetch(url, options);
    var code = response.getResponseCode();
    var body = response.getContentText();

    if (code === 429) {
      Logger.log(label + ' rate limited. Trying next option.');
      return null;
    }
    if (code !== 200) {
      Logger.log(label + ' error (HTTP ' + code + ').');
      return null;
    }

    var json = JSON.parse(body);
    var text = json.choices[0].message.content;
    text = text.replace(/```json\s*/g, '').replace(/```\s*/g, '').trim();
    Logger.log(label + ' succeeded.');
    return JSON.parse(text);
  } catch (e) {
    Logger.log(label + ' request or response processing failed.');
    return null;
  }
}


// ============================================================
// WRITE MONTHLY REPORT
// ============================================================
function writeMonthlyReport(ss, items, monthLabel) {
  var sheet = ss.getSheetByName(CONFIG.reportSheet);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.reportSheet);
  } else {
    sheet.clearContents();
    sheet.clearFormats();
    sheet.setConditionalFormatRules([]);
  }

  var row = 1;

  sheet.getRange(row, 1).setValue('COMPETITOR CONTENT REPORT: ' + monthLabel.toUpperCase());
  sheet.getRange(row, 1).setFontSize(14).setFontWeight('bold');
  row++;

  var datedCount = items.filter(function(i) { return i.date; }).length;
  var undatedCount = items.length - datedCount;
  sheet.getRange(row, 1).setValue(
    'Generated on ' + new Date().toLocaleDateString() +
    ' | ' + items.length + ' items found (' + datedCount + ' with confirmed dates, ' + undatedCount + ' undated)'
  );
  sheet.getRange(row, 1).setFontColor('#888888').setFontSize(9);
  row += 2;

  var headers = ['#', 'Competitor', 'Content type', 'Title', 'Publish date', 'URL', 'Summary', 'Data source'];
  sheet.getRange(row, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(row, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#1a1a1a')
    .setFontColor('#ffffff')
    .setFontSize(10)
    .setWrap(true);
  row++;

  if (items.length > 0) {
    items.sort(function(a, b) {
      // Dated items first, then undated
      if (a.date && !b.date) return -1;
      if (!a.date && b.date) return 1;
      if (a.competitor !== b.competitor) return a.competitor.localeCompare(b.competitor);
      if (a.date && b.date) return b.date - a.date;
      return 0;
    });

    var dataRows = [];
    for (var i = 0; i < items.length; i++) {
      var item = items[i];
      var dateStr = item.date
        ? Utilities.formatDate(item.date, Session.getScriptTimeZone(), 'MMM d, yyyy')
        : 'Unknown';

      dataRows.push([
        i + 1,
        item.competitor,
        item.contentType || 'Unknown',
        item.title,
        dateStr,
        item.url,
        item.summary || '(no summary available)',
        item.source || '',
      ]);
    }

    sheet.getRange(row, 1, dataRows.length, headers.length).setValues(dataRows);
    sheet.getRange(row, 1, dataRows.length, headers.length).setWrap(true).setVerticalAlignment('top');

    // Conditional formatting for content type
    var typeRange = sheet.getRange(row, 3, dataRows.length, 1);
    var rules = [];
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Blog').setBackground('#DBEAFE').setFontColor('#1E40AF').setRanges([typeRange]).build());
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Guide').setBackground('#FEE2E2').setFontColor('#991B1B').setRanges([typeRange]).build());
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Template').setBackground('#FEF3C7').setFontColor('#92400E').setRanges([typeRange]).build());
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Report').setBackground('#E0E7FF').setFontColor('#3730A3').setRanges([typeRange]).build());
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Webinar').setBackground('#F3E8FF').setFontColor('#6B21A8').setRanges([typeRange]).build());
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Lead magnet').setBackground('#D1FAE5').setFontColor('#065F46').setRanges([typeRange]).build());

    // Highlight undated rows
    var dateRange = sheet.getRange(row, 5, dataRows.length, 1);
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Unknown').setBackground('#FEF3C7').setFontColor('#92400E').setRanges([dateRange]).build());

    sheet.setConditionalFormatRules(rules);
    row += dataRows.length;
  } else {
    sheet.getRange(row, 1).setValue('No new content found for this month.');
    row++;
  }

  sheet.setColumnWidth(1, 40);
  sheet.setColumnWidth(2, 120);
  sheet.setColumnWidth(3, 100);
  sheet.setColumnWidth(4, 300);
  sheet.setColumnWidth(5, 100);
  sheet.setColumnWidth(6, 250);
  sheet.setColumnWidth(7, 350);
  sheet.setColumnWidth(8, 80);

  sheet.setFrozenRows(4);
}


// ============================================================
// APPEND TO ARCHIVE
// ============================================================
function appendToArchive(ss, items, monthLabel) {
  var sheet = ss.getSheetByName(CONFIG.archiveSheet);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.archiveSheet);
  }

  if (sheet.getLastRow() === 0) {
    var headers = ['Scan month', 'Competitor', 'Content type', 'Title', 'Publish date', 'URL', 'Summary'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#1a1a1a')
      .setFontColor('#ffffff');
  }

  if (items.length === 0) return;

  var startRow = sheet.getLastRow() + 1;
  var dataRows = [];

  for (var i = 0; i < items.length; i++) {
    var item = items[i];
    var dateStr = item.date
      ? Utilities.formatDate(item.date, Session.getScriptTimeZone(), 'MMM d, yyyy')
      : 'Unknown';

    dataRows.push([
      monthLabel,
      item.competitor,
      item.contentType || 'Unknown',
      item.title,
      dateStr,
      item.url,
      item.summary || '',
    ]);
  }

  sheet.getRange(startRow, 1, dataRows.length, 7).setValues(dataRows);
  sheet.getRange(startRow, 1, dataRows.length, 7).setWrap(true);
}


// ============================================================
// SLACK NOTIFICATION
// ============================================================
function sendSlackReport(items, monthLabel, sheetUrl) {
  if (!CONFIG.slackWebhookUrl || CONFIG.slackWebhookUrl.indexOf('hooks.slack.com') === -1) {
    Logger.log('Slack webhook not configured.');
    return;
  }

  var byCompetitor = {};
  for (var i = 0; i < items.length; i++) {
    var comp = items[i].competitor;
    if (!byCompetitor[comp]) byCompetitor[comp] = { blogs: 0, resources: 0, titles: [] };
    if (items[i].contentType === 'Blog') {
      byCompetitor[comp].blogs++;
    } else {
      byCompetitor[comp].resources++;
    }
    if (byCompetitor[comp].titles.length < 3) {
      byCompetitor[comp].titles.push(items[i].title);
    }
  }

  var blocks = [
    {
      type: 'header',
      text: { type: 'plain_text', text: 'Competitor content report: ' + monthLabel, emoji: true },
    },
    {
      type: 'section',
      text: { type: 'mrkdwn', text: 'Found *' + items.length + ' new content items* across ' + Object.keys(byCompetitor).length + ' competitors.' },
    },
    { type: 'divider' },
  ];

  var compNames = Object.keys(byCompetitor);
  for (var c = 0; c < compNames.length; c++) {
    var name = compNames[c];
    var data = byCompetitor[name];
    var countText = [];
    if (data.blogs > 0) countText.push(data.blogs + ' blog' + (data.blogs > 1 ? 's' : ''));
    if (data.resources > 0) countText.push(data.resources + ' resource' + (data.resources > 1 ? 's' : ''));

    var sampleTitles = data.titles.map(function(t) { return '\u2022 ' + t; }).join('\n');

    blocks.push({
      type: 'section',
      text: {
        type: 'mrkdwn',
        text: '*' + name + '* (' + countText.join(', ') + ')\n' + sampleTitles,
      },
    });
  }

  blocks.push({ type: 'divider' });
  blocks.push({
    type: 'section',
    text: { type: 'mrkdwn', text: '<' + sheetUrl + '|View full report in Google Sheets>' },
  });

  try {
    UrlFetchApp.fetch(CONFIG.slackWebhookUrl, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({ blocks: blocks }),
      muteHttpExceptions: true,
    });
    Logger.log('Slack notification sent.');
  } catch (e) {
    Logger.log('Slack notification failed.');
  }
}


// ============================================================
// HELPERS
// ============================================================
function fetchPage(url) {
  try {
    var response = UrlFetchApp.fetch(url, {
      muteHttpExceptions: true,
      followRedirects: true,
      headers: { 'User-Agent': 'Mozilla/5.0 (compatible; BlogTracker/1.0)' },
    });
    if (response.getResponseCode() !== 200) return null;
    return response.getContentText();
  } catch (e) {
    Logger.log('Fetch error for ' + url + ': ' + e.message);
    return null;
  }
}

function stripHTML(html) {
  return html.replace(/<[^>]+>/g, '').replace(/&amp;/g, '&').replace(/&nbsp;/g, ' ').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&#\d+;/g, '').replace(/&[a-z]+;/gi, ' ');
}

function cleanText(text) {
  return text.replace(/\s+/g, ' ').trim();
}

function getChildText(element, childName, namespace) {
  var child;
  try {
    child = namespace ? element.getChild(childName, namespace) : element.getChild(childName);
    if (!child) child = element.getChild(childName);
  } catch (e) {
    child = element.getChild(childName);
  }
  return child ? child.getText() : '';
}