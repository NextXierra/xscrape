import { chromium } from 'playwright-core';
import ExcelJS from 'exceljs';
import path from 'path';
import fs from 'fs';
import { fileURLToPath } from 'url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));

// Load exact 186 column headers in order
const COLUMNS_PATH = path.join(__dirname, 'columns_186.json');
const COLUMNS = JSON.parse(fs.readFileSync(COLUMNS_PATH, 'utf8'));

function parseArgs() {
  const args = process.argv.slice(2);
  let query = '';
  let targetCount = 100;
  let since = null;
  let until = null;
  let filter = 'live';
  let isExact = false;

  const positional = [];

  for (let i = 0; i < args.length; i++) {
    const arg = args[i];
    if (arg === '--since' || arg === '-s') {
      since = args[++i];
    } else if (arg === '--until' || arg === '-u') {
      until = args[++i];
    } else if (arg === '--filter' || arg === '-f') {
      filter = args[++i];
    } else if (arg === '--limit' || arg === '-l' || arg === '-n') {
      targetCount = parseInt(args[++i], 10) || 100;
    } else if (arg === '--query' || arg === '-q') {
      query = args[++i];
    } else if (arg === '--exact' || arg === '-e') {
      isExact = true;
    } else if (!arg.startsWith('-')) {
      positional.push(arg);
    }
  }

  // Parse positional tokens intelligently
  const queryTokens = [];
  for (const token of positional) {
    if (/^\d{4}-\d{2}-\d{2}$/.test(token)) {
      if (!since) since = token;
      else if (!until) until = token;
    } else if (/^\d+$/.test(token) && !args.includes('--limit') && !args.includes('-l') && !args.includes('-n')) {
      targetCount = parseInt(token, 10);
    } else {
      queryTokens.push(token);
    }
  }

  if (!query && queryTokens.length > 0) {
    let joined = queryTokens.join(' ').trim();
    // Detect if escaped quotes/backslashes were passed (e.g. from PowerShell \"phrase\")
    if (/^(\\+"|\\"|\\|'|")/.test(joined) && /(\\"|\\|'|")$/.test(joined)) {
      isExact = true;
      joined = joined.replace(/^(\\+"|\\"|\\|'|")+/, '').replace(/(\\"|\\|'|")+$/, '').trim();
    }
    query = joined;
  }

  if (isExact && query) {
    const unquoted = query.replace(/^"+|"+$/g, '').trim();
    query = `"${unquoted}"`;
  }

  const MAX_SAFE_LIMIT = 500;
  if (targetCount > MAX_SAFE_LIMIT) {
    targetCount = MAX_SAFE_LIMIT;
  }

  let finalQuery = query || 'genshin';
  if (since && !finalQuery.includes('since:')) {
    finalQuery += ` since:${since}`;
  }
  if (until && !finalQuery.includes('until:')) {
    finalQuery += ` until:${until}`;
  }

  return {
    rawQuery: query,
    query: finalQuery,
    targetCount,
    since,
    until,
    filter,
    isExact,
    cdpUrl: 'http://127.0.0.1:9222',
    scrollDelayMs: 1800,
    maxScrollAttempts: 400
  };
}

const CONFIG = parseArgs();

function renderProgress(current, total, width = 30) {
  const bounded = Math.min(current, total);
  const ratio = total > 0 ? bounded / total : 0;
  const filled = Math.round(ratio * width);
  const empty = width - filled;
  const bar = '='.repeat(filled) + '-'.repeat(empty);
  const percent = Math.round(ratio * 100).toString().padStart(3, ' ');
  process.stdout.write(`\r[${bar}] ${bounded}/${total} (${percent}%)`);
}

function toPyRepr(val) {
  if (val === null || val === undefined) return null;
  if (typeof val === 'boolean') return val ? 'True' : 'False';
  if (typeof val === 'number') return String(val);
  if (typeof val === 'string') {
    return "'" + val.replace(/\\/g, '\\\\').replace(/'/g, "\\'") + "'";
  }
  if (Array.isArray(val)) {
    return '[' + val.map(toPyRepr).join(', ') + ']';
  }
  if (typeof val === 'object') {
    const pairs = Object.entries(val).map(([k, v]) => `${toPyRepr(k)}: ${toPyRepr(v)}`);
    return '{' + pairs.join(', ') + '}';
  }
  return String(val);
}

function parseAuthor(userResult) {
  if (!userResult) return {};
  if (userResult.__typename === 'UserWithVisibilityResults' && userResult.user) {
    userResult = userResult.user;
  }
  const core = userResult.core || {};
  const legacy = userResult.legacy || {};
  const relCounts = userResult.relationship_counts || {};
  const tweetCounts = userResult.tweet_counts || {};
  const actionCounts = userResult.action_counts || {};
  const label = userResult.affiliates_highlighted_label?.label;

  const userName = core.screen_name || legacy.screen_name || '';
  const descUrls = legacy.entities?.description?.urls || [];
  const urlUrls = legacy.entities?.url?.urls || [];
  const pinnedIds = userResult.pinned_items?.tweet_ids_str || legacy.pinned_tweet_ids_str || [];

  return {
    type: 'user',
    userName: userName,
    url: userName ? `https://x.com/${userName}` : '',
    twitterUrl: userName ? `https://twitter.com/${userName}` : '',
    id: userResult.rest_id || legacy.id_str || '',
    name: core.name || legacy.name || '',
    isVerified: Boolean(userResult.verification?.verified ?? legacy.verified),
    isBlueVerified: Boolean(userResult.is_blue_verified),
    verifiedType: userResult.verification?.verified_type || null,
    profilePicture: userResult.avatar?.image_url || legacy.profile_image_url_https || '',
    coverPicture: userResult.banner?.image_url || legacy.profile_banner_url || null,
    description: null,
    location: userResult.location?.location || legacy.location || null,
    followers: Number(relCounts.followers ?? legacy.followers_count ?? 0),
    following: Number(relCounts.following ?? legacy.friends_count ?? 0),
    status: null,
    canDm: true,
    canMediaTag: userResult.media_permissions?.can_media_tag ?? true,
    createdAt: core.created_at || legacy.created_at || '',
    entities_description_urls: toPyRepr(descUrls),
    fastFollowersCount: Number(legacy.fast_followers_count || 0),
    favouritesCount: Number(actionCounts.favorites_count ?? legacy.favourites_count ?? 0),
    hasCustomTimelines: Boolean(legacy.has_custom_timelines ?? true),
    isTranslator: Boolean(legacy.is_translator ?? false),
    mediaCount: Number(tweetCounts.media_tweets ?? legacy.media_count ?? 0),
    statusesCount: Number(tweetCounts.tweets ?? legacy.statuses_count ?? 0),
    withheldInCountries: '[]',
    possiblySensitive: Boolean(userResult.possibly_sensitive ?? false),
    pinnedTweetIds: toPyRepr(pinnedIds),
    profile_bio_description: userResult.profile_bio?.description || legacy.description || null,
    isAutomated: Boolean(label?.user_label_type === 'AutomatedLabel'),
    automatedBy: label?.long_description_text ? label.long_description_text.replace(/^Automated by @?/, '') : null,
    profile_bio_entities_url_urls: urlUrls.length ? toPyRepr(urlUrls) : null,
    profile_bio_entities_description_urls: descUrls.length ? toPyRepr(descUrls) : null,
    profile_bio_entities_description_user_mentions: legacy.entities?.description?.user_mentions?.length ? toPyRepr(legacy.entities.description.user_mentions) : null,
    profile_bio_entities_description_hashtags: legacy.entities?.description?.hashtags?.length ? toPyRepr(legacy.entities.description.hashtags) : null,
    affiliatesHighlightedLabel_label_badge_url: label?.badge?.url || null,
    affiliatesHighlightedLabel_label_description: label?.description || null,
    affiliatesHighlightedLabel_label_long_description_entities: label?.long_description_entities ? toPyRepr(label.long_description_entities) : null,
    affiliatesHighlightedLabel_label_long_description_text: label?.long_description_text || null,
    affiliatesHighlightedLabel_label_user_label_type: label?.user_label_type || null
  };
}

function extractExact186Row(rawTweet) {
  let tweet = rawTweet;
  if (tweet.__typename === 'TweetWithVisibilityResults' && tweet.tweet) {
    tweet = tweet.tweet;
  }
  const legacy = tweet.legacy;
  if (!legacy || !legacy.id_str) return null;

  let userResult = tweet.core?.user_results?.result;
  const author = parseAuthor(userResult);

  const cleanSrc = tweet.source ? tweet.source.replace(/<[^>]+>/g, '') : '';
  const idStr = legacy.id_str;
  const userName = author.userName;

  // Quoted tweet
  let quotedResult = tweet.quoted_status_result?.result;
  let quotedTweet = null;
  if (quotedResult) {
    if (quotedResult.__typename === 'TweetWithVisibilityResults' && quotedResult.tweet) {
      quotedResult = quotedResult.tweet;
    }
    if (quotedResult.legacy) {
      const qLegacy = quotedResult.legacy;
      const qAuthor = parseAuthor(quotedResult.core?.user_results?.result);
      const qSrc = quotedResult.source ? quotedResult.source.replace(/<[^>]+>/g, '') : '';

      let qqResult = quotedResult.quoted_status_result?.result;
      let qqTweet = null;
      if (qqResult) {
        if (qqResult.__typename === 'TweetWithVisibilityResults' && qqResult.tweet) {
          qqResult = qqResult.tweet;
        }
        if (qqResult.legacy) {
          qqTweet = {
            type: 'tweet',
            id: qqResult.legacy.id_str,
            source: qqResult.source ? qqResult.source.replace(/<[^>]+>/g, '') : '',
            retweetCount: Number(qqResult.legacy.retweet_count || 0),
            replyCount: Number(qqResult.legacy.reply_count || 0),
            likeCount: Number(qqResult.legacy.favorite_count || 0),
            quoteCount: Number(qqResult.legacy.quote_count || 0),
            viewCount: qqResult.views?.count ? Number(qqResult.views.count) : 0,
            bookmarkCount: Number(qqResult.legacy.bookmark_count || 0),
            isReply: Boolean(qqResult.legacy.in_reply_to_status_id_str),
            displayTextRange: toPyRepr(qqResult.legacy.display_text_range || []),
            isLimitedReply: Boolean(qqResult.legacy.limited_actions)
          };
        }
      }

      quotedTweet = {
        type: 'tweet',
        id: qLegacy.id_str,
        url: qAuthor.userName ? `https://x.com/${qAuthor.userName}/status/${qLegacy.id_str}` : '',
        twitterUrl: qAuthor.userName ? `https://twitter.com/${qAuthor.userName}/status/${qLegacy.id_str}` : '',
        text: qLegacy.full_text || '',
        source: qSrc,
        retweetCount: Number(qLegacy.retweet_count || 0),
        replyCount: Number(qLegacy.reply_count || 0),
        likeCount: Number(qLegacy.favorite_count || 0),
        quoteCount: Number(qLegacy.quote_count || 0),
        viewCount: quotedResult.views?.count ? Number(quotedResult.views.count) : 0,
        createdAt: qLegacy.created_at || '',
        lang: qLegacy.lang || '',
        bookmarkCount: Number(qLegacy.bookmark_count || 0),
        isReply: Boolean(qLegacy.in_reply_to_status_id_str),
        conversationId: qLegacy.conversation_id_str || qLegacy.id_str,
        displayTextRange: toPyRepr(qLegacy.display_text_range || []),
        author: qAuthor,
        extendedEntities_media: qLegacy.extended_entities?.media ? toPyRepr(qLegacy.extended_entities.media) : null,
        isLimitedReply: Boolean(qLegacy.limited_actions),
        entities_hashtags: qLegacy.entities?.hashtags?.length ? toPyRepr(qLegacy.entities.hashtags) : null,
        entities_symbols: qLegacy.entities?.symbols?.length ? toPyRepr(qLegacy.entities.symbols) : null,
        entities_urls: qLegacy.entities?.urls?.length ? toPyRepr(qLegacy.entities.urls) : null,
        entities_user_mentions: qLegacy.entities?.user_mentions?.length ? toPyRepr(qLegacy.entities.user_mentions) : null,
        quoted_tweet: qqTweet
      };
    }
  }

  const card = tweet.card;
  const cardPlatform = card?.legacy?.card_platform?.platform;
  const place = legacy.place;

  // Exact row structure mapping
  const row = {};
  for (const col of COLUMNS) {
    row[col] = null;
  }

  row.type = 'tweet';
  row.id = idStr;
  row.url = userName ? `https://x.com/${userName}/status/${idStr}` : '';
  row.twitterUrl = userName ? `https://twitter.com/${userName}/status/${idStr}` : '';
  row.text = legacy.full_text || '';
  row.source = cleanSrc;
  row.retweetCount = Number(legacy.retweet_count || 0);
  row.replyCount = Number(legacy.reply_count || 0);
  row.likeCount = Number(legacy.favorite_count || 0);
  row.quoteCount = Number(legacy.quote_count || 0);
  row.viewCount = tweet.views?.count ? Number(tweet.views.count) : 0;
  row.createdAt = legacy.created_at || '';
  row.lang = legacy.lang || '';
  row.bookmarkCount = Number(legacy.bookmark_count || 0);
  row.isReply = Boolean(legacy.in_reply_to_status_id_str);
  row.inReplyToId = legacy.in_reply_to_status_id_str || null;
  row.conversationId = legacy.conversation_id_str || idStr;
  row.displayTextRange = toPyRepr(legacy.display_text_range || []);
  row.inReplyToUserId = legacy.in_reply_to_user_id_str || null;
  row.inReplyToUsername = legacy.in_reply_to_screen_name || null;
  row.isLimitedReply = Boolean(legacy.limited_actions);

  // Author fields
  for (const [k, v] of Object.entries(author)) {
    const key = `author_${k}`;
    if (key in row) {
      row[key] = v;
    }
  }

  // Tweet entities
  if (legacy.entities?.user_mentions?.length) row.entities_user_mentions = toPyRepr(legacy.entities.user_mentions);
  if (legacy.extended_entities?.media?.length) row.extendedEntities_media = toPyRepr(legacy.extended_entities.media);
  if (legacy.entities?.hashtags?.length) row.entities_hashtags = toPyRepr(legacy.entities.hashtags);
  if (legacy.entities?.symbols?.length) row.entities_symbols = toPyRepr(legacy.entities.symbols);
  if (legacy.entities?.urls?.length) row.entities_urls = toPyRepr(legacy.entities.urls);
  if (legacy.entities?.timestamps?.length) row.entities_timestamps = toPyRepr(legacy.entities.timestamps);

  // Quoted tweet fields
  if (quotedTweet) {
    row.quoted_tweet_type = quotedTweet.type;
    row.quoted_tweet_id = quotedTweet.id;
    row.quoted_tweet_url = quotedTweet.url;
    row.quoted_tweet_twitterUrl = quotedTweet.twitterUrl;
    row.quoted_tweet_text = quotedTweet.text;
    row.quoted_tweet_source = quotedTweet.source;
    row.quoted_tweet_retweetCount = quotedTweet.retweetCount;
    row.quoted_tweet_replyCount = quotedTweet.replyCount;
    row.quoted_tweet_likeCount = quotedTweet.likeCount;
    row.quoted_tweet_quoteCount = quotedTweet.quoteCount;
    row.quoted_tweet_viewCount = quotedTweet.viewCount;
    row.quoted_tweet_createdAt = quotedTweet.createdAt;
    row.quoted_tweet_lang = quotedTweet.lang;
    row.quoted_tweet_bookmarkCount = quotedTweet.bookmarkCount;
    row.quoted_tweet_isReply = quotedTweet.isReply;
    row.quoted_tweet_conversationId = quotedTweet.conversationId;
    row.quoted_tweet_displayTextRange = quotedTweet.displayTextRange;
    row.quoted_tweet_extendedEntities_media = quotedTweet.extendedEntities_media;
    row.quoted_tweet_isLimitedReply = quotedTweet.isLimitedReply;
    row.quoted_tweet_entities_hashtags = quotedTweet.entities_hashtags;
    row.quoted_tweet_entities_symbols = quotedTweet.entities_symbols;
    row.quoted_tweet_entities_urls = quotedTweet.entities_urls;
    row.quoted_tweet_entities_user_mentions = quotedTweet.entities_user_mentions;

    if (quotedTweet.author) {
      for (const [ak, av] of Object.entries(quotedTweet.author)) {
        const key = `quoted_tweet_author_${ak}`;
        if (key in row) {
          row[key] = av;
        }
      }
    }

    if (quotedTweet.quoted_tweet) {
      const qq = quotedTweet.quoted_tweet;
      row.quoted_tweet_quoted_tweet_type = qq.type;
      row.quoted_tweet_quoted_tweet_id = qq.id;
      row.quoted_tweet_quoted_tweet_source = qq.source;
      row.quoted_tweet_quoted_tweet_retweetCount = qq.retweetCount;
      row.quoted_tweet_quoted_tweet_replyCount = qq.replyCount;
      row.quoted_tweet_quoted_tweet_likeCount = qq.likeCount;
      row.quoted_tweet_quoted_tweet_quoteCount = qq.quoteCount;
      row.quoted_tweet_quoted_tweet_viewCount = qq.viewCount;
      row.quoted_tweet_quoted_tweet_bookmarkCount = qq.bookmarkCount;
      row.quoted_tweet_quoted_tweet_isReply = qq.isReply;
      row.quoted_tweet_quoted_tweet_displayTextRange = qq.displayTextRange;
      row.quoted_tweet_quoted_tweet_isLimitedReply = qq.isLimitedReply;
    }
  }

  // Card fields
  if (card) {
    row.card_name = card.legacy?.name || null;
    row.card_url = card.legacy?.url || null;
    if (card.legacy?.binding_values) row.card_binding_values = toPyRepr(card.legacy.binding_values);
    if (cardPlatform?.audience?.name) row.card_card_platform_platform_audience_name = cardPlatform.audience.name;
    if (cardPlatform?.device?.name) row.card_card_platform_platform_device_name = cardPlatform.device.name;
    if (cardPlatform?.device?.version) row.card_card_platform_platform_device_version = cardPlatform.device.version;
    if (card.legacy?.user_refs_results) row.card_user_refs_results = toPyRepr(card.legacy.user_refs_results);
  }

  // Place fields
  if (place) {
    row.place_id = place.id || null;
    row.place_name = place.name || null;
    row.place_full_name = place.full_name || null;
    row.place_country = place.country || null;
    row.place_country_code = place.country_code || null;
    row.place_place_type = place.place_type || null;
    if (place.bounding_box) {
      row.place_bounding_box_polygon_type = place.bounding_box.type || null;
      row.place_bounding_box_polygon_coordinates = toPyRepr(place.bounding_box.coordinates || []);
    }
  }

  return row;
}

function findTweetsInObject(obj, found = []) {
  if (!obj || typeof obj !== 'object') return found;

  if (obj.__typename === 'Tweet' || (obj.__typename === 'TweetWithVisibilityResults' && obj.tweet)) {
    found.push(obj);
    return found;
  }

  if (Array.isArray(obj)) {
    for (const item of obj) {
      findTweetsInObject(item, found);
    }
  } else {
    for (const key of Object.keys(obj)) {
      findTweetsInObject(obj[key], found);
    }
  }

  return found;
}

const OUTPUT_DIR = path.join(__dirname, 'output');
if (!fs.existsSync(OUTPUT_DIR)) {
  fs.mkdirSync(OUTPUT_DIR, { recursive: true });
}

const HISTORY_FILE = path.join(__dirname, '.scraped_ids.json');

async function loadExistingScrapedIds() {
  const ids = new Set();
  
  if (fs.existsSync(HISTORY_FILE)) {
    try {
      const data = JSON.parse(fs.readFileSync(HISTORY_FILE, 'utf8'));
      if (Array.isArray(data)) {
        for (const id of data) ids.add(String(id));
      }
    } catch {}
  }

  // Scan folder output dan direktori root
  const searchDirs = [OUTPUT_DIR, __dirname];
  for (const dir of searchDirs) {
    if (!fs.existsSync(dir)) continue;
    const files = fs.readdirSync(dir).filter(f => f.startsWith('tweets_') && f.endsWith('.xlsx'));
    for (const file of files) {
      try {
        const wb = new ExcelJS.Workbook();
        await wb.xlsx.readFile(path.join(dir, file));
        const ws = wb.getWorksheet(1);
        if (ws) {
          ws.eachRow((row, rowNumber) => {
            if (rowNumber === 1) return;
            const val = row.getCell(2).value;
            if (val) ids.add(String(val));
          });
        }
      } catch {}
    }
  }

  try {
    fs.writeFileSync(HISTORY_FILE, JSON.stringify(Array.from(ids)), 'utf8');
  } catch {}

  return ids;
}

function updateScrapedHistory(newIds) {
  try {
    let ids = [];
    if (fs.existsSync(HISTORY_FILE)) {
      ids = JSON.parse(fs.readFileSync(HISTORY_FILE, 'utf8')) || [];
    }
    const set = new Set(ids);
    for (const id of newIds) set.add(String(id));
    fs.writeFileSync(HISTORY_FILE, JSON.stringify(Array.from(set)), 'utf8');
  } catch {}
}

async function saveToExcel(tweets, outputPath) {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('Sheet1');

  worksheet.columns = COLUMNS.map(c => ({ header: c, key: c }));

  for (const tweet of tweets) {
    worksheet.addRow(tweet);
  }

  await workbook.xlsx.writeFile(outputPath);
}

async function run() {
  const searchUrl = `https://x.com/search?q=${encodeURIComponent(CONFIG.query)}&f=${CONFIG.filter}`;
  const rangeStr = (CONFIG.since || CONFIG.until) ? ` [${CONFIG.since || ''} .. ${CONFIG.until || ''}]` : '';
  const displayQuery = (CONFIG.rawQuery || CONFIG.query).replace(/^"+|"+$/g, '');
  console.log(`[xscrape] "${displayQuery}"${rangeStr} | Target: ${CONFIG.targetCount}`);

  let browser;
  try {
    const previouslyScrapedIds = await loadExistingScrapedIds();

    browser = await chromium.connectOverCDP(CONFIG.cdpUrl);
    const context = browser.contexts()[0];
    if (!context) {
      throw new Error('Koneksi Edge gagal: browser context tidak ditemukan.');
    }

    let page = context.pages().find(p => p.url().includes('x.com'));
    if (!page) {
      page = await context.newPage();
    }

    const tweetMap = new Map();

    page.on('response', async (response) => {
      const url = response.url();
      if (url.includes('/graphql/') && (url.includes('SearchTimeline') || url.includes('TweetDetail') || url.includes('Timeline'))) {
        try {
          const contentType = response.headers()['content-type'] || '';
          if (contentType.includes('json')) {
            const json = await response.json();
            const foundRawTweets = findTweetsInObject(json);

            for (const raw of foundRawTweets) {
              const parsed = extractExact186Row(raw);
              if (parsed && parsed.id) {
                if (previouslyScrapedIds.has(parsed.id) || tweetMap.has(parsed.id)) {
                  continue;
                }
                tweetMap.set(parsed.id, parsed);
                renderProgress(tweetMap.size, CONFIG.targetCount);
              }
            }
          }
        } catch {}
      }
    });

    renderProgress(0, CONFIG.targetCount);

    await page.goto(searchUrl, { waitUntil: 'domcontentloaded' });
    await page.waitForTimeout(2500);

    let scrollAttempts = 0;
    let lastCount = 0;
    let consecutiveStale = 0;
    let endOfTimeline = false;

    while (tweetMap.size < CONFIG.targetCount && scrollAttempts < CONFIG.maxScrollAttempts) {
      scrollAttempts++;

      await page.evaluate(() => {
        window.scrollBy({
          top: window.innerHeight * 1.5,
          behavior: 'smooth'
        });
      }).catch(() => {});

      await page.waitForTimeout(CONFIG.scrollDelayMs);

      if (tweetMap.size === lastCount) {
        consecutiveStale++;

        const pageStatus = await page.evaluate(() => {
          const text = document.body ? document.body.innerText : '';
          const hasEmpty = text.includes('No results for') ||
                           text.includes('No search results') ||
                           text.includes('These results are up to date') ||
                           text.includes('No Tweets found') ||
                           text.includes('Tidak ada hasil untuk');
          const hasError = text.includes('Something went wrong. Try reloading.') ||
                           text.includes('Retry');
          const isAtBottom = (window.innerHeight + window.scrollY) >= (document.documentElement.scrollHeight - 120);
          return { hasEmpty, hasError, isAtBottom };
        }).catch(() => ({ hasEmpty: false, hasError: false, isAtBottom: false }));

        if (pageStatus.hasEmpty) {
          endOfTimeline = true;
          break;
        }

        // Coba scroll kejut / re-trigger jika sempat macet di ronde ke-3
        if (consecutiveStale === 3) {
          await page.evaluate(() => window.scrollBy(0, -400)).catch(() => {});
          await page.waitForTimeout(600);
          await page.evaluate(() => window.scrollBy(0, 800)).catch(() => {});
        }

        // Jika sudah mentok di ujung bawah halaman dan tidak ada tweet baru
        if (pageStatus.isAtBottom && consecutiveStale >= 4) {
          endOfTimeline = true;
          break;
        }

        // Jika sudah 7 ronde scroll berturut-turut tanpa tweet baru
        if (consecutiveStale >= 7) {
          endOfTimeline = true;
          break;
        }
      } else {
        consecutiveStale = 0;
        lastCount = tweetMap.size;
      }
    }

    renderProgress(Math.min(tweetMap.size, CONFIG.targetCount), CONFIG.targetCount);
    process.stdout.write('\n');

    const tweetsArray = Array.from(tweetMap.values()).slice(0, CONFIG.targetCount);

    if (tweetsArray.length === 0) {
      console.log(`[xscrape] Selesai: 0 tweet ditemukan.\n`);
      return;
    }

    const sanitizedQuery = (CONFIG.rawQuery || CONFIG.query).replace(/[^a-zA-Z0-9]/g, '_').slice(0, 30);
    const dateTag = CONFIG.since ? `_${CONFIG.since}` : '';
    const timestamp = new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19);
    const filename = `tweets_${sanitizedQuery}${dateTag}_${tweetsArray.length}_${timestamp}.xlsx`;
    const outputPath = path.join(OUTPUT_DIR, filename);

    await saveToExcel(tweetsArray, outputPath);
    updateScrapedHistory(tweetsArray.map(t => t.id));

    console.log(`[xscrape] Selesai: ${tweetsArray.length} tweet tersimpan -> output/${filename}\n`);

  } catch (err) {
    if (err.message && err.message.includes('ECONNREFUSED')) {
      console.error('\n[xscrape] Error: Microsoft Edge belum aktif. Jalankan `node launch.js` terlebih dahulu.\n');
    } else {
      console.error(`\n[xscrape] Error: ${err.message}\n`);
    }
  } finally {
    if (browser) {
      await browser.close();
    }
  }
}

run();
