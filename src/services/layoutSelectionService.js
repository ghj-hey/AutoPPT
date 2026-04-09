const DEFAULT_LAYOUTS_BY_TYPE = {
  cover: "cover_formal_v1",
  summary_cards: "summary_grid_v1",
  table_analysis: "table_compare_v1",
  process_flow: "process_three_lane_v1",
  bullet_columns: "bullet_dual_v1",
  image_story: "image_split_v1",
  action_plan: "action_dashboard_v1",
  key_takeaways: "takeaway_cards_v1",
};

function friendlyTemplateLabel(id = "") {
  const value = String(id || "").trim();
  const map = {
    cover_formal_v1: "正式封面",
    cover_clean_v1: "简洁封面",
    summary_grid_v1: "摘要四卡",
    summary_dense_v1: "高密度摘要",
    summary_spread_v1: "横向摘要铺陈",
    summary_bands_v1: "摘要条带",
    table_compare_v1: "对比表格",
    table_split_v1: "图表拆分",
    table_visual_v1: "图表混排",
    table_dashboard_v1: "看板表格",
    table_dense_v1: "高密度表格",
    table_sidecallout_v1: "侧栏结论表格",
    table_picture_v1: "图片佐证表格",
    table_matrix_v1: "矩阵分析表格",
    process_three_lane_v1: "三段流程",
    process_bridge_v1: "桥接流程",
    process_cards_v1: "流程卡片",
    process_ladder_v1: "阶梯流程",
    bullet_dual_v1: "双栏要点",
    bullet_triple_v1: "三栏要点",
    bullet_staggered_v1: "错落要点",
    bullet_masonry_v1: "瀑布要点",
    image_split_v1: "图文分栏",
    image_focus_v1: "主图聚焦",
    image_storyboard_v1: "图文故事板",
    image_gallery_v1: "图片画廊",
    action_dashboard_v1: "行动看板",
    action_stacked_v1: "堆叠行动卡",
    action_timeline_v1: "时间推进",
    takeaway_cards_v1: "结论卡片",
    takeaway_wall_v1: "结论墙",
  };
  return map[value] || value;
}

function friendlyPageTypeLabel(type = "") {
  const value = String(type || "").trim();
  const map = {
    cover: "封面页",
    summary_cards: "摘要页",
    table_analysis: "表格分析页",
    process_flow: "流程方法页",
    bullet_columns: "分栏要点页",
    image_story: "图文展示页",
    action_plan: "行动计划页",
    key_takeaways: "结论收束页",
  };
  return map[value] || "页面";
}

function friendlyLayoutSetLabel(id = "") {
  const value = String(id || "").trim();
  const map = {
    bank_finance_default: "标准金融汇报",
    bank_finance_dense: "高密度汇报",
    bank_finance_visual: "图文强化",
    bank_finance_highlight: "重点高亮",
    bank_finance_boardroom: "正式工作会",
    bank_finance_reporting: "综合材料型",
    bank_finance_dynamic: "动态组合型",
  };
  return map[value] || value;
}

function sanitizeReadableLabel(text = "", fallback = "") {
  const value = String(text || "").trim();
  if (!value || /[?]{2,}|[�]/.test(value)) {
    return fallback || "";
  }
  return value;
}


function unique(list) {
  return [...new Set((list || []).filter(Boolean))];
}

function stableHash(value = "") {
  let hash = 0;
  const text = String(value || "");
  for (let index = 0; index < text.length; index += 1) {
    hash = (hash * 31 + text.charCodeAt(index)) % 2147483647;
  }
  return Math.abs(hash);
}

const TEMPLATE_SKELETON_REGISTRY = {
  table_analysis: {
    dense: "data_wall",
    dashboard: "data_wall",
    matrix: "matrix_grid",
    compare: "compare_split",
    sidecallout: "compare_split",
    split: "compare_split",
    visual: "media_panel",
    picture: "media_panel",
    highlight: "highlight_focus",
    stack: "stacked_story",
  },
};

function templateVariantFamily(variant = "", pageType = "") {
  const value = String(variant || "").toLowerCase();
  const page = String(pageType || "").toLowerCase();
  const registryFamily = TEMPLATE_SKELETON_REGISTRY[page]?.[value];
  if (registryFamily) return registryFamily;
  if (!value) return "general";
  if (["dense", "dashboard", "matrix"].includes(value)) return "dense";
  if (["split", "compare", "sidecallout"].includes(value)) return "split";
  if (["visual", "picture", "focus", "gallery", "storyboard"].includes(value)) return "visual";
  if (["highlight", "callout", "badge", "ribbon"].includes(value)) return "split";
  if (["stack", "wall", "closing"].includes(value)) return "stack";
  if (["bridge", "cards", "ladder", "three_lane"].includes(value)) return "process";
  if (["spread", "grid", "mosaic", "dense_grid", "bands"].includes(value)) return "summary";
  if (["timeline", "stacked", "action", "roadmap"].includes(value)) return "action";
  if (["triple", "dual", "staggered", "masonry"].includes(value)) return "bullet";
  return value;
}

function inferContentMode(slide = {}) {
  const textLength = [
    slide.headline || "",
    ...(slide.cards || []).map((item) => item.body || item.detail || ""),
    ...(slide.insights || []).map((item) => item.body || ""),
    ...(slide.columns || []).flatMap((item) => item.bullets || []),
    ...(slide.stages || []).map((item) => item.detail || ""),
    ...(slide.steps || []).map((item) => item.detail || ""),
    ...(slide.takeaways || []).map((item) => item.body || ""),
    slide.footer || "",
  ]
    .join(" ")
    .length;
  const metricCount = slide.metrics?.length || 0;
  const cardCount = slide.cards?.length || 0;
  const insightCount = slide.insights?.length || 0;
  const rowCount = slide.table?.rowCount || slide.table?.rows?.length || 0;
  const colCount = slide.table?.colCount || slide.table?.header?.length || 0;
  const stageCount = slide.stages?.length || 0;
  const stepCount = slide.steps?.length || 0;
  const takeawaysCount = slide.takeaways?.length || 0;
  const hasImage = Boolean(slide.image?.path);
  const hasScreenshots = Boolean(slide.screenshots?.length);

  const denseScore =
    (rowCount >= 5 ? 2 : 0) +
    (colCount >= 4 ? 1 : 0) +
    (metricCount >= 4 ? 1 : 0) +
    (cardCount >= 4 ? 1 : 0) +
    (insightCount >= 3 ? 1 : 0) +
    (stageCount >= 4 ? 1 : 0) +
    (stepCount >= 4 ? 1 : 0) +
    (takeawaysCount >= 4 ? 1 : 0) +
    (textLength >= 280 ? 2 : 0);
  const visualScore = (hasImage ? 2 : 0) + (hasScreenshots ? 2 : 0) + (cardCount <= 3 ? 1 : 0) + (textLength <= 240 ? 1 : 0);
  const sparseScore = (textLength <= 180 ? 2 : 0) + (metricCount <= 2 ? 1 : 0) + (cardCount <= 2 ? 1 : 0);
  const processScore = (stageCount >= 3 ? 2 : 0) + (stepCount >= 3 ? 2 : 0) + (textLength <= 320 ? 1 : 0);

  const scores = [
    { mode: "dense", score: denseScore },
    { mode: "visual", score: visualScore },
    { mode: "sparse", score: sparseScore },
    { mode: "process", score: processScore },
  ].sort((left, right) => right.score - left.score);

  if ((scores[0]?.score || 0) <= 0) return "balanced";
  if ((scores[0]?.score || 0) - (scores[1]?.score || 0) <= 1 && (scores[0]?.mode === "dense" || scores[0]?.mode === "visual")) {
    return "balanced";
  }
  return scores[0].mode;
}

function styleVariantPreferences(referenceStyleProfile = {}, slideType = "") {
  const family = String(referenceStyleProfile?.styleFamily || "").toLowerCase();
  const density = String(referenceStyleProfile?.densityBias || "").toLowerCase();
  const tablePreference = String(referenceStyleProfile?.tablePreference || "").toLowerCase();
  const pageRhythm = String(referenceStyleProfile?.pageRhythm || "").toLowerCase();
  const imagePlacement = String(referenceStyleProfile?.imagePlacement || "").toLowerCase();
  const cardStyle = String(referenceStyleProfile?.cardStyle || "").toLowerCase();
  const layoutBias = Array.isArray(referenceStyleProfile?.templateBias) ? referenceStyleProfile.templateBias : [];
  const profilePreferred = Array.isArray(referenceStyleProfile?.preferredVariants) ? referenceStyleProfile.preferredVariants : [];
  const preferred = [];

  if (family === "dense-report" || density === "high") {
    if (slideType === "summary_cards") preferred.push("dense_grid", "mosaic", "grid", "bands", "spread");
    if (slideType === "table_analysis") preferred.push("dense", "dashboard", "matrix", "compare", "stack", "picture", "highlight", "sidecallout", "visual", "split");
    if (slideType === "process_flow") preferred.push("ladder", "bridge", "cards", "three_lane");
    if (slideType === "bullet_columns") preferred.push("masonry", "triple", "staggered", "dual");
    if (slideType === "image_story") preferred.push("storyboard", "gallery", "focus", "split");
    if (slideType === "action_plan") preferred.push("dashboard", "matrix", "stacked", "timeline");
    if (slideType === "key_takeaways") preferred.push("wall", "closing", "cards");
  } else if (family === "visual-report") {
    if (slideType === "summary_cards") preferred.push("mosaic", "spread", "grid", "dense_grid", "bands");
    if (slideType === "table_analysis") preferred.push("visual", "picture", "compare", "sidecallout", "highlight", "split", "dashboard", "matrix", "stack", "dense");
    if (slideType === "process_flow") preferred.push("bridge", "cards", "ladder", "three_lane");
    if (slideType === "bullet_columns") preferred.push("staggered", "dual", "triple", "masonry");
    if (slideType === "image_story") preferred.push("storyboard", "focus", "gallery", "split");
    if (slideType === "action_plan") preferred.push("dashboard", "matrix", "timeline", "stacked");
    if (slideType === "key_takeaways") preferred.push("closing", "cards", "wall");
  } else if (family === "boardroom-report") {
    if (slideType === "summary_cards") preferred.push("spread", "grid", "bands", "dense_grid", "mosaic");
    if (slideType === "table_analysis") preferred.push("split", "compare", "highlight", "sidecallout", "visual", "picture", "dashboard", "matrix", "stack", "dense");
    if (slideType === "process_flow") preferred.push("cards", "bridge", "three_lane", "ladder");
    if (slideType === "bullet_columns") preferred.push("dual", "staggered", "triple", "masonry");
    if (slideType === "image_story") preferred.push("split", "gallery", "focus", "storyboard");
    if (slideType === "action_plan") preferred.push("timeline", "matrix", "dashboard", "stacked");
    if (slideType === "key_takeaways") preferred.push("closing", "cards", "wall");
  }

  if (tablePreference) {
    if (slideType === "table_analysis") {
      if (tablePreference === "dense") preferred.unshift("dense", "dashboard", "matrix");
      if (tablePreference === "compare") preferred.unshift("compare", "sidecallout", "highlight", "split");
      if (tablePreference === "picture") preferred.unshift("picture", "visual", "storyboard");
      if (tablePreference === "dashboard") preferred.unshift("dashboard", "matrix", "compare");
    }
    if (slideType === "summary_cards" && tablePreference === "dashboard") {
      preferred.unshift("dense_grid", "mosaic");
    }
  }

  if (pageRhythm) {
    if (pageRhythm === "dense") {
      if (slideType === "summary_cards") preferred.unshift("dense_grid", "mosaic");
      if (slideType === "table_analysis") preferred.unshift("dense", "dashboard", "matrix");
      if (slideType === "bullet_columns") preferred.unshift("triple", "masonry");
      if (slideType === "action_plan") preferred.unshift("dashboard", "matrix");
    } else if (pageRhythm === "visual") {
      if (slideType === "summary_cards") preferred.unshift("spread", "mosaic", "bands");
      if (slideType === "table_analysis") preferred.unshift("visual", "picture");
      if (slideType === "image_story") preferred.unshift("storyboard", "gallery", "focus");
    }
  }

  if (imagePlacement) {
    if (slideType === "image_story") {
      if (imagePlacement === "gallery") preferred.unshift("gallery", "storyboard");
      if (imagePlacement === "hero") preferred.unshift("focus", "split");
      if (imagePlacement === "split") preferred.unshift("split", "storyboard");
    }
    if (slideType === "table_analysis" && imagePlacement === "gallery") {
      preferred.unshift("picture", "visual");
    }
  }

  if (cardStyle) {
    if (slideType === "summary_cards") {
      if (cardStyle === "dashboard-card") preferred.unshift("dense_grid", "mosaic");
      if (cardStyle === "ribbon-card") preferred.unshift("bands", "grid");
      if (cardStyle === "soft-card") preferred.unshift("spread", "mosaic");
    }
    if (slideType === "bullet_columns" && cardStyle === "dashboard-card") {
      preferred.unshift("masonry", "triple");
    }
    if (slideType === "action_plan" && cardStyle === "ribbon-card") {
      preferred.unshift("timeline", "stacked");
    }
  }

  if (profilePreferred.length) {
    preferred.unshift(...profilePreferred.map((item) => String(item || "").toLowerCase()).filter(Boolean));
  }

  if (layoutBias.length) {
    const bias = layoutBias.map((item) => String(item || "").toLowerCase()).filter(Boolean);
    if (slideType === "summary_cards") {
      if (bias.some((item) => item.includes("dense"))) preferred.unshift("dense_grid", "mosaic");
      if (bias.some((item) => item.includes("spread"))) preferred.unshift("spread", "grid");
    }
    if (slideType === "table_analysis") {
      if (bias.some((item) => item.includes("dense"))) preferred.unshift("dense", "dashboard");
      if (bias.some((item) => item.includes("split"))) preferred.unshift("split", "compare", "sidecallout");
      if (bias.some((item) => item.includes("visual"))) preferred.unshift("visual", "picture", "sidecallout");
      if (bias.some((item) => item.includes("matrix") || item.includes("table"))) preferred.unshift("matrix", "dense");
      if (bias.some((item) => item.includes("dashboard"))) preferred.unshift("dashboard", "matrix");
    }
    if (slideType === "image_story") {
      if (bias.some((item) => item.includes("image") || item.includes("visual"))) preferred.unshift("storyboard", "gallery", "focus");
    }
    if (slideType === "action_plan") {
      if (bias.some((item) => item.includes("timeline"))) preferred.unshift("timeline", "stacked");
      if (bias.some((item) => item.includes("dashboard"))) preferred.unshift("dashboard", "matrix");
    }
  }

  return unique(preferred);
}

function layoutBiasVariants(layoutBias = "", familyHint = "", contentBalance = "") {
  const bias = String(layoutBias || "").toLowerCase();
  const family = String(familyHint || "").toLowerCase();
  const balance = String(contentBalance || "").toLowerCase();
  const preferred = [];

  if (bias.includes("picture") || bias.includes("image") || bias.includes("gallery") || bias.includes("focus")) {
    preferred.push("picture", "visual", "split", "gallery", "focus", "storyboard");
  }
  if (bias.includes("highlight") || bias.includes("callout")) {
    preferred.push("highlight", "compare", "sidecallout", "split", "visual");
  }
  if (bias.includes("dashboard") || bias.includes("matrix")) {
    preferred.push("dashboard", "matrix", "dense", "compare", "stack");
  }
  if (bias.includes("dense") || bias.includes("spread")) {
    preferred.push("dense", "dashboard", "matrix", "compare");
  }
  if (bias.includes("bridge") || bias.includes("flow") || bias.includes("cards")) {
    preferred.push("bridge", "cards", "ladder", "three_lane");
  }
  if (bias.includes("timeline") || bias.includes("roadmap")) {
    preferred.push("timeline", "stacked", "dashboard", "matrix");
  }
  if (bias.includes("wall") || bias.includes("closing")) {
    preferred.push("wall", "closing", "cards");
  }
  if (bias.includes("mixed") || balance === "mixed") {
    preferred.push("split", "compare", "visual", "picture", "dashboard", "matrix", "highlight");
  }

  if (family === "visual") {
    preferred.push("visual", "picture", "gallery", "focus", "storyboard", "split");
  } else if (family === "process") {
    preferred.push("bridge", "cards", "ladder", "three_lane");
  } else if (family === "dense") {
    preferred.push("dense", "dashboard", "matrix", "compare");
  } else if (family === "mixed") {
    preferred.push("split", "compare", "visual", "picture", "dashboard", "matrix");
  } else if (family === "stack") {
    preferred.push("stack", "wall", "closing", "cards");
  } else if (family === "action") {
    preferred.push("timeline", "stacked", "dashboard", "matrix");
  } else if (family === "summary") {
    preferred.push("spread", "grid", "bands", "dense_grid", "mosaic");
  }

  if (balance === "image-heavy") preferred.unshift("picture", "visual", "gallery");
  if (balance === "table-heavy" || balance === "chart-heavy") preferred.unshift("dashboard", "matrix", "dense");
  if (balance === "mixed") preferred.unshift("split", "compare", "visual", "picture", "dashboard", "matrix");
  if (balance === "text-heavy") preferred.unshift("staggered", "masonry", "dual", "triple");

  return unique(preferred);
}

function tableModeVariants(tableMode = "", pageParity = 0, contentMode = "") {
  const mode = String(tableMode || "").toLowerCase();
  const parity = Number(pageParity) % 2;
  const balance = String(contentMode || "").toLowerCase();
  const preferred = [];

  if (["dashboard", "dense", "matrix"].includes(mode)) {
    preferred.push(
      ...(parity === 0 ? ["dashboard", "dense", "matrix"] : ["dense", "matrix", "dashboard"]),
      "compare",
      "sidecallout",
      "highlight",
      "split",
      "visual",
    );
  } else if (["compare", "sidecallout", "split"].includes(mode)) {
    preferred.push(
      ...(parity === 0 ? ["compare", "sidecallout", "split"] : ["sidecallout", "compare", "split"]),
      "highlight",
      "visual",
      "picture",
      "dashboard",
      "matrix",
      "dense",
    );
  } else if (["visual", "picture", "storyboard", "gallery", "focus"].includes(mode)) {
    preferred.push(
      ...(parity === 0 ? ["visual", "picture", "storyboard"] : ["picture", "visual", "storyboard"]),
      "gallery",
      "focus",
      "split",
      "compare",
      "sidecallout",
      "highlight",
      "dashboard",
      "matrix",
      "dense",
    );
  } else if (["stack", "highlight", "cards", "closing"].includes(mode)) {
    preferred.push(
      ...(parity === 0 ? ["stack", "highlight", "cards"] : ["highlight", "stack", "cards"]),
      "closing",
      "sidecallout",
      "compare",
      "visual",
      "split",
    );
  } else if (mode === "mixed") {
    preferred.push(
      ...(parity === 0 ? ["split", "compare", "visual"] : ["compare", "split", "visual"]),
      "picture",
      "dashboard",
      "matrix",
      "sidecallout",
      "highlight",
      "dense",
      "stack",
    );
  }

  if (balance === "image-heavy") preferred.unshift("visual", "picture", "gallery");
  if (balance === "table-heavy" || balance === "chart-heavy") preferred.unshift("dashboard", "matrix", "dense");
  if (balance === "mixed") preferred.unshift("split", "compare", "visual", "picture", "dashboard", "matrix");
  if (balance === "text-heavy") preferred.unshift("stack", "highlight", "cards");

  return unique(preferred);
}

function templateIdsByType(layoutLibrary, slideType) {
  return Object.entries(layoutLibrary?.templates || {})
    .filter(([, template]) => template.pageType === slideType)
    .map(([id, template]) => ({
      id,
      ...template,
      family: template.family || templateVariantFamily(template.variant, slideType),
    }));
}

function chooseCandidateTemplate(templatesByType, variants, previousTemplate = "", seed = "", usageCounts = {}, previousFamily = "", recentFamilies = [], slideType = "") {
  const orderedIds = unique(
    variants
      .map((variant) => templatesByType.find((template) => template.variant === variant)?.id || "")
      .filter(Boolean),
  );

  if (!orderedIds.length) return "";
  const scored = orderedIds.map((id, index) => {
    const template = templatesByType.find((item) => item.id === id) || {};
    const family = template.family || templateVariantFamily(template.variant, slideType);
    const used = Number(usageCounts[id] || 0);
    const repeatPenalty = id === previousTemplate ? (slideType === "table_analysis" ? 2200 : 1000) : 0;
    const familyPenalty = previousFamily && family === previousFamily ? (slideType === "table_analysis" ? 1100 : 360) : 0;
    const recentPenalty = (recentFamilies || [])
      .slice(-4)
      .reverse()
      .findIndex((item) => item === family);
    const recentFamilyPenalty = recentPenalty >= 0 ? (4 - recentPenalty) * (slideType === "table_analysis" ? 260 : 120) : 0;
    const seedBonus = (stableHash(`${seed}|${id}`) % 97) / 100;
    return {
      id,
      score: index * 100 + used * 24 + repeatPenalty + familyPenalty + recentFamilyPenalty + seedBonus,
    };
  });
  scored.sort((left, right) => left.score - right.score);
  return scored[0]?.id || orderedIds[0] || "";
}

function preferredVariantsForSlide(slide, referenceStyleProfile = {}) {
  const contentMode = String(slide.contentMode || inferContentMode(slide) || "balanced").toLowerCase();
  const density = String(slide.density || "medium").toLowerCase();
  const layoutBias = String(slide.layoutBias || "").toLowerCase();
  const layoutTier = String(slide.layoutTier || "").toLowerCase();
  const tableMode = String(slide.tableMode || "").toLowerCase();
  const familyHint = String(slide.familyHint || "").toLowerCase();
  const contentBalance = String(slide.contentBalance || "").toLowerCase();
  const slidePreferredFamilies = unique(
    (Array.isArray(slide.preferredFamilies) ? slide.preferredFamilies : [])
      .map((item) => String(item || "").toLowerCase())
      .filter(Boolean),
  );
  const preferredVariantHints = unique(
    (Array.isArray(referenceStyleProfile?.preferredVariants) ? referenceStyleProfile.preferredVariants : [])
      .map((item) => String(item || "").toLowerCase())
      .filter(Boolean),
  );
  const pageNo = Number(slide.page || 0);
  const pageParity = pageNo % 2;
  const pageCycle = pageNo % 3;
  const hasImage = Boolean(slide.image?.path);
  const hasScreenshots = Boolean(slide.screenshots?.length);
  const hasBars = Boolean((slide.bars || []).filter((item) => Number(item?.value || 0) > 0).length);
  const rowCount = slide.table?.rowCount || slide.table?.rows?.length || 0;
  const colCount = slide.table?.colCount || slide.table?.header?.length || 0;
  const denseTableSignal = rowCount >= 5 || colCount >= 5 || (rowCount >= 4 && colCount >= 6);
  const sparseTableSignal = rowCount <= 3 && colCount <= 4;
  const cardCount = slide.cards?.length || 0;
  const insightCount = slide.insights?.length || 0;
  const stageCount = slide.stages?.length || 0;
  const stepCount = slide.steps?.length || 0;
  const takeawaysCount = slide.takeaways?.length || 0;
  const stylePrefs = styleVariantPreferences(referenceStyleProfile, slide.type);
  const biasPrefs = layoutBiasVariants(layoutBias, familyHint, contentBalance);
  const tableModePrefs = slide.type === "table_analysis" ? tableModeVariants(tableMode, pageParity, contentMode) : [];
  const cyclicalPrefs = (() => {
    if (slide.type === "table_analysis") {
      if (pageCycle === 0) return ["dashboard", "matrix", "dense", "compare", "sidecallout", "split", "highlight", "visual", "picture", "stack"];
      if (pageCycle === 1) return ["compare", "sidecallout", "highlight", "split", "matrix", "visual", "picture", "dashboard", "dense", "stack"];
      return ["visual", "picture", "stack", "highlight", "compare", "sidecallout", "dashboard", "matrix", "dense", "split"];
    }
    if (slide.type === "image_story") {
      if (pageCycle === 0) return ["storyboard", "gallery", "focus", "split"];
      if (pageCycle === 1) return ["focus", "split", "gallery", "storyboard"];
      return ["split", "gallery", "focus", "storyboard"];
    }
    if (slide.type === "action_plan") {
      if (pageCycle === 0) return ["dashboard", "matrix", "stacked", "timeline"];
      if (pageCycle === 1) return ["timeline", "stacked", "dashboard", "matrix"];
      return ["stacked", "dashboard", "timeline", "matrix"];
    }
    if (slide.type === "summary_cards") {
      if (pageCycle === 0) return ["dense_grid", "mosaic", "grid", "bands", "spread"];
      if (pageCycle === 1) return ["mosaic", "spread", "bands", "grid", "dense_grid"];
      return ["spread", "bands", "grid", "mosaic", "dense_grid"];
    }
    return [];
  })();

  let variants = [];
  switch (slide.type) {
    case "summary_cards":
      if (layoutTier === "dense") {
        variants = pageParity === 0 ? ["dense_grid", "mosaic", "grid", "bands", "spread"] : ["mosaic", "dense_grid", "grid", "spread", "bands"];
      } else if (layoutTier === "visual") {
        variants = pageParity === 0 ? ["mosaic", "spread", "bands", "grid", "dense_grid"] : ["spread", "mosaic", "bands", "grid", "dense_grid"];
      } else if (layoutTier === "spacious") {
        variants = pageParity === 0 ? ["spread", "grid", "bands", "mosaic", "dense_grid"] : ["grid", "spread", "bands", "mosaic", "dense_grid"];
      } else if (contentMode === "sparse" || (density === "low" && cardCount >= 4)) {
        variants = pageParity === 0 ? ["spread", "mosaic", "grid", "bands", "dense_grid"] : ["bands", "spread", "grid", "mosaic", "dense_grid"];
      } else if (contentMode === "dense" || density === "high" || cardCount >= 5) {
        variants = pageParity === 0 ? ["dense_grid", "mosaic", "spread", "grid", "bands"] : ["mosaic", "dense_grid", "grid", "spread", "bands"];
      } else if (contentMode === "visual") {
        variants = pageParity === 0 ? ["spread", "mosaic", "bands", "grid", "dense_grid"] : ["mosaic", "spread", "bands", "grid", "dense_grid"];
      } else {
        variants = pageParity === 0 ? ["grid", "spread", "bands", "mosaic", "dense_grid"] : ["spread", "grid", "bands", "dense_grid", "mosaic"];
      }
      break;
    case "table_analysis":
      if (contentBalance === "mixed") {
        variants = pageParity === 0
          ? ["split", "compare", "visual", "picture", "sidecallout", "dashboard", "dense", "matrix", "highlight", "stack"]
          : ["compare", "split", "visual", "picture", "sidecallout", "dashboard", "dense", "matrix", "highlight", "stack"];
      }
      if (variants.length) break;
      if (tableMode === "dashboard") {
        variants = ["dashboard", "dense", "matrix", "compare", "sidecallout", "highlight", "split", "visual", "picture"];
      } else if (tableMode === "dense") {
        variants = ["dense", "dashboard", "matrix", "compare", "sidecallout", "highlight", "split", "visual", "picture"];
      } else if (tableMode === "matrix") {
        variants = ["matrix", "dashboard", "dense", "compare", "sidecallout", "highlight", "split", "visual", "picture"];
      } else if (tableMode === "compare") {
        variants = ["compare", "sidecallout", "split", "highlight", "visual", "picture", "dashboard", "matrix", "dense"];
      } else if (tableMode === "sidecallout") {
        variants = ["sidecallout", "compare", "split", "highlight", "visual", "picture", "dashboard", "matrix", "dense"];
      } else if (tableMode === "visual") {
        variants = ["visual", "picture", "storyboard", "gallery", "split", "compare", "sidecallout", "highlight", "dashboard", "matrix"];
      } else if (tableMode === "picture") {
        variants = ["picture", "visual", "storyboard", "gallery", "split", "compare", "sidecallout", "highlight", "dashboard", "matrix"];
      } else if (tableMode === "stack") {
        variants = ["stack", "highlight", "cards", "closing", "split", "compare", "visual"];
      }
      if (variants.length) break;
      if (layoutTier === "dense") {
        variants = ["dashboard", "matrix", "dense", "compare", "split", "highlight", "sidecallout", "visual"];
      } else if (layoutTier === "visual") {
        variants = ["visual", "picture", "storyboard", "gallery", "split", "compare", "sidecallout", "highlight", "dashboard", "matrix"];
      } else if (layoutTier === "compare") {
        variants = ["compare", "split", "sidecallout", "highlight", "dashboard", "visual", "matrix", "dense"];
      } else if (layoutTier === "highlight") {
        variants = ["highlight", "sidecallout", "compare", "split", "visual", "dashboard", "matrix", "dense"];
      } else if (layoutTier === "spacious") {
        variants = ["picture", "visual", "split", "compare", "sidecallout", "highlight", "dashboard", "matrix", "dense"];
      } else if (layoutBias.includes("picture") || layoutBias.includes("visual")) {
        variants = ["picture", "visual", "split", "compare", "sidecallout", "dashboard", "highlight", "matrix", "dense"];
      } else if (layoutBias.includes("highlight") || layoutBias.includes("callout")) {
        variants = ["highlight", "compare", "sidecallout", "split", "dashboard", "visual", "matrix", "dense"];
      } else if (layoutBias.includes("dashboard") || layoutBias.includes("matrix") || layoutBias.includes("dense")) {
        variants = ["dashboard", "matrix", "dense", "compare", "highlight", "sidecallout", "split", "visual"];
      } else if (layoutBias.includes("compare") || layoutBias.includes("split")) {
        variants = ["compare", "split", "sidecallout", "highlight", "dashboard", "visual", "matrix", "dense"];
      }
      if (variants.length) break;
      if (hasBars && rowCount <= 4 && colCount <= 4) {
        variants = pageParity === 0
          ? ["compare", "sidecallout", "highlight", "split", "visual", "picture", "dashboard", "dense", "matrix"]
          : ["sidecallout", "compare", "highlight", "split", "visual", "picture", "dashboard", "dense", "matrix"];
      } else if (denseTableSignal) {
        variants = pageParity === 0
          ? ["dashboard", "dense", "matrix", "compare", "sidecallout", "highlight", "split", "visual", "picture"]
          : ["compare", "dashboard", "dense", "matrix", "sidecallout", "highlight", "split", "visual", "picture"];
      } else if ((contentMode === "visual" || hasImage || hasScreenshots) && !denseTableSignal) {
        variants = pageParity === 0
          ? ["compare", "visual", "sidecallout", "highlight", "dashboard", "split", "matrix", "dense", "picture"]
          : ["sidecallout", "compare", "visual", "highlight", "dashboard", "split", "matrix", "dense", "picture"];
      } else if (contentMode === "dense" || rowCount >= 7 || colCount >= 5 || (density === "high" && rowCount >= 5)) {
        variants = pageParity === 0
          ? ["dashboard", "matrix", "dense", "compare", "highlight", "sidecallout", "split", "visual", "stack"]
          : ["matrix", "dense", "dashboard", "compare", "highlight", "sidecallout", "split", "visual", "stack"];
      } else if (contentMode === "sparse" || sparseTableSignal) {
        variants = pageParity === 0
          ? ["highlight", "stack", "compare", "split", "visual", "picture", "sidecallout", "dashboard", "dense", "matrix"]
          : ["stack", "highlight", "compare", "visual", "split", "picture", "sidecallout", "dashboard", "matrix", "dense"];
      } else if (insightCount >= 3) {
        variants = ["highlight", "sidecallout", "compare", "stack", "matrix", "picture", "visual", "split", "dense"];
      } else if (rowCount >= 5 && colCount >= 4) {
        variants = ["matrix", "dashboard", "dense", "split", "compare", "highlight", "picture", "sidecallout", "visual", "stack"];
      } else {
        variants = ["split", "compare", "highlight", "picture", "visual", "stack", "dashboard", "sidecallout", "dense", "matrix"];
      }
      variants = [...cyclicalPrefs, ...variants];
      break;
    case "process_flow":
      if (layoutTier === "dense" || layoutTier === "process") {
        variants = ["ladder", "bridge", "cards", "three_lane"];
      } else if (layoutTier === "visual") {
        variants = ["bridge", "cards", "three_lane", "ladder"];
      } else if (layoutBias.includes("bridge") || layoutBias.includes("cards") || layoutBias.includes("ladder")) {
        variants = ["bridge", "cards", "ladder", "three_lane"];
      } else if (layoutBias.includes("timeline")) {
        variants = ["timeline", "stacked", "dashboard", "matrix"];
      }
      if (variants.length) break;
      if (contentMode === "dense" || stageCount >= 5 || density === "high") {
        variants = pageParity === 0 ? ["ladder", "bridge", "cards", "three_lane"] : ["bridge", "ladder", "cards", "three_lane"];
      } else if (contentMode === "visual" || stepCount >= 4) {
        variants = pageParity === 0 ? ["bridge", "cards", "ladder", "three_lane"] : ["cards", "bridge", "ladder", "three_lane"];
      } else if (stageCount <= 4 && (slide.notes?.length || 0) <= 2) {
        variants = pageParity === 0 ? ["cards", "bridge", "three_lane", "ladder"] : ["bridge", "cards", "three_lane", "ladder"];
      } else {
        variants = pageParity === 0 ? ["three_lane", "bridge", "cards", "ladder"] : ["bridge", "cards", "three_lane", "ladder"];
      }
      break;
    case "bullet_columns":
      if (layoutTier === "dense") {
        variants = ["masonry", "triple", "staggered", "dual"];
      } else if (layoutTier === "spacious") {
        variants = ["dual", "staggered", "triple", "masonry"];
      } else if (layoutBias.includes("masonry") || layoutBias.includes("triple") || layoutBias.includes("staggered")) {
        variants = ["masonry", "triple", "staggered", "dual"];
      } else if (layoutBias.includes("dual")) {
        variants = ["dual", "staggered", "triple", "masonry"];
      }
      if (variants.length) break;
      if (contentMode === "dense" || ((slide.columns?.length || 0) >= 3 && cardCount >= 2 && density === "high")) variants = ["masonry", "triple", "staggered", "dual"];
      else if (contentMode === "sparse" || (cardCount >= 2 && rowCount <= 0)) variants = ["staggered", "dual", "triple", "masonry"];
      else if ((slide.columns?.length || 0) >= 3 || density === "high") variants = ["triple", "masonry", "staggered", "dual"];
      else variants = ["dual", "staggered", "triple", "masonry"];
      break;
    case "image_story":
      if (layoutTier === "visual") {
        variants = ["storyboard", "gallery", "focus", "split"];
      } else if (layoutTier === "spacious") {
        variants = ["focus", "gallery", "storyboard", "split"];
      } else if (layoutBias.includes("split")) {
        variants = ["split", "gallery", "focus", "storyboard"];
      } else if (layoutBias.includes("gallery")) {
        variants = ["gallery", "storyboard", "focus", "split"];
      } else if (layoutBias.includes("focus")) {
        variants = ["focus", "gallery", "storyboard", "split"];
      } else if (layoutBias.includes("storyboard")) {
        variants = ["storyboard", "gallery", "focus", "split"];
      }
      if (variants.length) break;
      if (hasImage && (slide.callouts?.length || 0) <= 1 && (slide.bullets?.length || 0) <= 1) variants = ["split", "focus", "gallery", "storyboard"];
      else if (contentMode === "visual" || (hasImage && (slide.callouts?.length || 0) >= 2)) variants = ["storyboard", "gallery", "split", "focus"];
      else if (hasImage && (slide.callouts?.length || 0) >= 3 && density !== "high") variants = ["gallery", "storyboard", "split", "focus"];
      else if ((slide.callouts?.length || 0) >= 3 && density === "high") variants = ["gallery", "storyboard", "focus", "split"];
      else variants = ["split", "gallery", "focus", "storyboard"];
      variants = [...cyclicalPrefs, ...variants];
      break;
    case "action_plan":
      if (layoutTier === "dense" || layoutTier === "action") {
        variants = ["dashboard", "matrix", "timeline", "stacked"];
      } else if (layoutTier === "visual") {
        variants = ["timeline", "dashboard", "matrix", "stacked"];
      } else if (layoutBias.includes("timeline") || layoutBias.includes("stacked")) {
        variants = ["timeline", "stacked", "dashboard", "matrix"];
      } else if (layoutBias.includes("dashboard") || layoutBias.includes("matrix")) {
        variants = ["dashboard", "matrix", "timeline", "stacked"];
      }
      if (variants.length) break;
      if (hasScreenshots && stepCount >= 2) variants = ["dashboard", "matrix", "stacked", "timeline"];
      else if (contentMode === "visual") variants = ["dashboard", "stacked", "matrix", "timeline"];
      else if (contentMode === "dense" || density === "high") variants = ["stacked", "dashboard", "matrix", "timeline"];
      else if (contentMode === "sparse" || stepCount <= 3) variants = ["stacked", "dashboard", "timeline", "matrix"];
      else variants = ["timeline", "dashboard", "matrix", "stacked"];
      variants = [...cyclicalPrefs, ...variants];
      break;
    case "key_takeaways":
      if (layoutTier === "spacious" || layoutTier === "summary") {
        variants = ["wall", "closing", "cards"];
      } else if (layoutBias.includes("wall") || layoutBias.includes("closing")) {
        variants = ["wall", "closing", "cards"];
      }
      if (variants.length) break;
      if (contentMode === "dense" || takeawaysCount >= 4) variants = ["wall", "closing", "cards"];
      else if (contentMode === "sparse" || takeawaysCount <= 2) variants = ["closing", "cards", "wall"];
      else variants = ["cards", "closing", "wall"];
      break;
    case "cover":
    default:
      variants = [];
  }

  return unique([...preferredVariantHints, ...slidePreferredFamilies, ...tableModePrefs, ...biasPrefs, ...stylePrefs, ...variants]);
}

function resolveTemplateId(slide, layoutLibrary, slideSelections = {}, context = {}) {
  const templatesByType = templateIdsByType(layoutLibrary, slide.type);
  const previousTemplate = context.previousTemplateByType?.[slide.type] || "";
  const previousFamily = context.previousFamilyByType?.[slide.type] || "";
  const usageCounts = context.usageCounts || {};
  const recentFamilies = context.recentFamilies || [];
  const selection = slideSelections[String(slide.page)] || "";
  const pageParity = Number(slide.page || 0) % 2;
  if (selection) return selection;
  if (slide.templateId) return slide.templateId;

  const preferred = preferredVariantsForSlide(slide, context.referenceStyleProfile || {});
  const seed = `${slide.page}|${slide.title || ""}|${slide.type}|${slide.contentMode || slide.density || ""}`;
  const recommended = chooseCandidateTemplate(templatesByType, preferred, previousTemplate, seed, usageCounts, previousFamily, recentFamilies, slide.type);
  if (recommended) {
    const recommendedTemplate = templatesByType.find((item) => item.id === recommended) || {};
    const recommendedVariant = recommendedTemplate.variant || "";
    const rowCount = slide.table?.rowCount || slide.table?.rows?.length || 0;
    const colCount = slide.table?.colCount || slide.table?.header?.length || 0;
    const hasDenseTable = rowCount >= 6 || colCount >= 5 || (rowCount >= 4 && colCount >= 4);
    const hasScreenshots = Boolean(slide.screenshots?.length);
    const screenshotCount = slide.screenshots?.length || 0;
    const stepCount = slide.steps?.length || 0;
    const calloutCount = slide.callouts?.length || 0;
    const hasImage = Boolean(slide.image?.path);

    const byVariant = (variant) => templatesByType.find((item) => item.variant === variant)?.id || "";

    if (slide.type === "table_analysis") {
      const previousFamilyKey = String(previousFamily || "").toLowerCase();
      const previousLooksDense = previousFamilyKey === "data_wall" || previousFamilyKey === "dense";
      const previousLooksSplit = previousFamilyKey === "compare_split" || previousFamilyKey === "split";
      const previousLooksMatrix = previousFamilyKey === "matrix_grid";
      const previousLooksFocus = previousFamilyKey === "highlight_focus";
      const previousLooksStack = previousFamilyKey === "stacked_story";
      const previousVariant = String(templatesByType.find((item) => item.id === previousTemplate)?.variant || "").toLowerCase();
      const hasNarrativeSignal = (slide.insights?.length || 0) >= 2 || (slide.metrics?.length || 0) >= 2 || Boolean(slide.table?.highlight);
      if (hasImage || hasScreenshots) {
        if (rowCount <= 4 && colCount <= 4) {
          if (previousVariant === "picture" || previousVariant === "visual") {
            return pageParity === 0
              ? byVariant("sidecallout") || byVariant("compare") || byVariant("split") || recommended
              : byVariant("compare") || byVariant("sidecallout") || byVariant("split") || recommended;
          }
          return pageParity === 0
            ? byVariant("picture") || byVariant("visual") || byVariant("compare") || byVariant("sidecallout") || recommended
            : byVariant("visual") || byVariant("picture") || byVariant("sidecallout") || byVariant("compare") || recommended;
        }
        return pageParity === 0
          ? byVariant("compare") || byVariant("sidecallout") || byVariant("visual") || byVariant("picture") || recommended
          : byVariant("sidecallout") || byVariant("compare") || byVariant("visual") || byVariant("picture") || recommended;
      }
      if ((slide.bars || []).length && rowCount <= 4 && colCount <= 4) {
        if (previousVariant === "compare" || previousVariant === "sidecallout") {
          return pageParity === 0
            ? byVariant("highlight") || byVariant("visual") || byVariant("sidecallout") || recommended
            : byVariant("visual") || byVariant("highlight") || byVariant("compare") || recommended;
        }
        return pageParity === 0
          ? byVariant("compare") || byVariant("sidecallout") || byVariant("highlight") || recommended
          : byVariant("sidecallout") || byVariant("compare") || byVariant("highlight") || recommended;
      }
      if (hasDenseTable) {
        if (previousLooksDense || previousLooksSplit || recommendedVariant === "dashboard" || recommendedVariant === "dense") {
          return pageParity === 0
            ? byVariant("matrix") || byVariant("compare") || byVariant("sidecallout") || byVariant("split") || recommended
            : byVariant("compare") || byVariant("matrix") || byVariant("sidecallout") || byVariant("picture") || recommended;
        }
        if (previousLooksMatrix || recommendedVariant === "matrix") {
          return pageParity === 0
            ? byVariant("compare") || byVariant("dashboard") || byVariant("dense") || byVariant("sidecallout") || recommended
            : byVariant("sidecallout") || byVariant("compare") || byVariant("dashboard") || byVariant("dense") || recommended;
        }
        return pageParity === 0
          ? byVariant("dashboard") || byVariant("dense") || byVariant("matrix") || byVariant("compare") || recommended
          : byVariant("compare") || byVariant("sidecallout") || byVariant("visual") || byVariant("picture") || recommended;
      }
      if (rowCount <= 3 && colCount <= 4) {
        if ((previousLooksFocus || previousLooksStack) && !hasImage && !hasScreenshots) {
          return pageParity === 0
            ? byVariant("compare") || byVariant("visual") || byVariant("picture") || recommended
            : byVariant("visual") || byVariant("compare") || byVariant("picture") || recommended;
        }
        if (hasNarrativeSignal && !hasImage && !hasScreenshots) {
          return pageParity === 0
            ? byVariant("highlight") || byVariant("stack") || byVariant("compare") || recommended
            : byVariant("stack") || byVariant("highlight") || byVariant("visual") || recommended;
        }
        if (previousVariant === "compare" || previousVariant === "sidecallout" || previousVariant === "split") {
          return pageParity === 0
            ? byVariant("visual") || byVariant("highlight") || byVariant("picture") || recommended
            : byVariant("sidecallout") || byVariant("highlight") || byVariant("visual") || recommended;
        }
        return pageParity === 0
          ? byVariant("compare") || byVariant("sidecallout") || byVariant("visual") || recommended
          : byVariant("visual") || byVariant("picture") || byVariant("compare") || recommended;
      }
    }
    if (slide.type === "action_plan" && hasScreenshots && screenshotCount >= 1 && stepCount >= 2 && recommendedVariant === "timeline") {
      return byVariant("dashboard") || byVariant("matrix") || byVariant("stacked") || recommended;
    }
    if (slide.type === "action_plan" && !hasScreenshots && stepCount > 0 && stepCount <= 3 && recommendedVariant === "timeline") {
      return byVariant("stacked") || byVariant("dashboard") || recommended;
    }
    if (slide.type === "image_story" && hasImage && calloutCount <= 1 && recommendedVariant === "focus") {
      return byVariant("split") || byVariant("gallery") || recommended;
    }
    if (slide.type === "image_story" && hasImage && calloutCount >= 2 && (slide.textBlocks || []).join(" ").length > 160 && recommendedVariant === "focus") {
      return byVariant("storyboard") || byVariant("split") || byVariant("gallery") || recommended;
    }
    return recommended;
  }

  return (
    layoutLibrary?.set?.[slide.type] ||
    layoutLibrary?.defaultsByType?.[slide.type] ||
    DEFAULT_LAYOUTS_BY_TYPE[slide.type] ||
    templatesByType[0]?.id ||
    ""
  );
}

function applyTemplatesToOutline(outline, layoutLibrary, slideSelections = {}, context = {}) {
  const previousTemplateByType = {};
  const previousFamilyByType = {};
  const usageCounts = {};
  const recentFamilies = [];
  return {
    ...outline,
    slides: (outline.slides || []).map((slide) => {
      const templateId = resolveTemplateId(slide, layoutLibrary, slideSelections, {
        ...context,
        previousTemplateByType,
        previousFamilyByType,
        usageCounts,
        recentFamilies,
      });
      previousTemplateByType[slide.type] = templateId;
      const family = layoutLibrary?.templates?.[templateId]?.family || templateVariantFamily(layoutLibrary?.templates?.[templateId]?.variant);
      const resolvedFamily = layoutLibrary?.templates?.[templateId]?.family || templateVariantFamily(layoutLibrary?.templates?.[templateId]?.variant, slide.type);
      previousFamilyByType[slide.type] = resolvedFamily;
      recentFamilies.push(resolvedFamily);
      while (recentFamilies.length > 4) recentFamilies.shift();
      usageCounts[templateId] = (usageCounts[templateId] || 0) + 1;
      return {
        ...slide,
        templateId,
      };
    }),
  };
}

function buildLayoutOptions(layoutLibrary, outline, currentSelection = {}, referenceStyleProfile = null) {
  const setOptions = Object.entries(layoutLibrary.sets || {}).map(([id, templates]) => ({
    id,
    displayName: sanitizeReadableLabel(layoutLibrary.setMetadata?.[id]?.displayName, friendlyLayoutSetLabel(id)),
    description: sanitizeReadableLabel(layoutLibrary.setMetadata?.[id]?.description, "用于适配不同金融汇报页面密度和风格。"),
    templates,
  }));

  const slideLayouts = [];
  const previousTemplateByType = {};
  const previousFamilyByType = {};
  const usageCounts = {};
  const recentFamilies = [];
  (outline?.slides || []).forEach((slide) => {
    const options = templateIdsByType(layoutLibrary, slide.type).map((template) => ({
      id: template.id,
      displayName: sanitizeReadableLabel(template.displayName, friendlyTemplateLabel(template.id)),
      label: friendlyTemplateLabel(template.id),
      description: sanitizeReadableLabel(template.description, "用于拉开不同页面的版式差异。"),
      variant: template.variant || "",
      family: template.family || templateVariantFamily(template.variant, slide.type),
    }));

    const recommendedTemplate = resolveTemplateId(slide, layoutLibrary, {}, {
      previousTemplateByType,
      previousFamilyByType,
      referenceStyleProfile,
      usageCounts,
      recentFamilies,
    });
    previousTemplateByType[slide.type] = recommendedTemplate;
    const recommendedFamily =
      layoutLibrary?.templates?.[recommendedTemplate]?.family ||
      templateVariantFamily(layoutLibrary?.templates?.[recommendedTemplate]?.variant, slide.type);
    previousFamilyByType[slide.type] = recommendedFamily;
    recentFamilies.push(recommendedFamily);
    while (recentFamilies.length > 4) recentFamilies.shift();
    usageCounts[recommendedTemplate] = (usageCounts[recommendedTemplate] || 0) + 1;

    const currentTemplate =
      currentSelection[String(slide.page)] ||
      slide.templateId ||
      recommendedTemplate ||
      layoutLibrary.set?.[slide.type] ||
      layoutLibrary.defaultsByType?.[slide.type] ||
      options[0]?.id ||
      "";

    slideLayouts.push({
      page: slide.page,
      type: slide.type,
      typeLabel: friendlyPageTypeLabel(slide.type),
      title: slide.title,
      hasImage: Boolean(slide.image?.path),
      hasScreenshots: Boolean(slide.screenshots?.length),
      tableRowCount: slide.table?.rowCount || slide.table?.rows?.length || 0,
      tableColCount: slide.table?.colCount || slide.table?.header?.length || 0,
      metricCount: slide.metrics?.length || 0,
      cardCount: slide.cards?.length || 0,
      insightCount: slide.insights?.length || 0,
      stageCount: slide.stages?.length || 0,
      stepCount: slide.steps?.length || 0,
      takeawayCount: slide.takeaways?.length || 0,
      layoutBias: slide.layoutBias || "",
      layoutTier: slide.layoutTier || "",
      familyHint: slide.familyHint || "",
      preferredFamilies: slide.preferredFamilies || [],
      contentBalance: slide.contentBalance || "",
      iconVariety: slide.iconVariety || "",
      spacingBias: slide.spacingBias || "",
      currentTemplate,
      recommendedTemplate,
      contentMode: slide.contentMode || inferContentMode(slide),
      density: slide.density || "medium",
      reason: [referenceStyleProfile?.styleFamily || "", slide.layoutBias || "", slide.layoutTier || "", slide.contentBalance || ""].filter(Boolean).join(" 璺?"),
      options,
    });
  });

  const initialSelection = Object.fromEntries(
    slideLayouts.map((slide) => [String(slide.page), slide.recommendedTemplate || slide.currentTemplate]),
  );

  return {
    layoutLibraryPath: layoutLibrary.path || "",
    activeSet: layoutLibrary.setName,
    defaultSet: layoutLibrary.defaultSet,
    setOptions,
    slideLayouts,
    initialSelection,
    referenceStyleProfile,
  };
}

module.exports = {
  DEFAULT_LAYOUTS_BY_TYPE,
  applyTemplatesToOutline,
  buildLayoutOptions,
  inferContentMode,
  templateVariantFamily,
  resolveTemplateId,
};

