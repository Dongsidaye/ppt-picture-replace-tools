"""PowerPoint-only source identity rules for destructive picture groups."""
from __future__ import annotations


def _replace_once(text: str, old: str, new: str, label: str) -> str:
    # The generator calls this before branding; the isolated test may receive
    # an already branded snapshot. These two comments are the only variants.
    if old not in text and "for WPS compatibility." in old:
        old = old.replace("for WPS compatibility.", "for PowerPoint compatibility.")
    count = text.count(old)
    if count != 1:
        raise ValueError(f"{label}: expected one target, found {count}")
    return text.replace(old, new, 1)


def prepare_identity(text: str) -> str:
    """Require a verified original-media SHA-256 identity before grouping."""
    text = _replace_once(text,
        '      const sourceKey = item.meta && item.meta.contentFp ? "m:" + item.meta.contentFp : "p:" + inst.uid;',
        '      // Persisted link metadata can outlive native Change Picture.\n'
        '      // Provisional rows stay independent until current media is verified.\n'
        '      const sourceKey = "p:" + inst.uid;',
        "unverified provisional sources")
    text = _replace_once(text,
        '          contentFp: item.meta ? item.meta.contentFp : "",',
        '          contentFp: "",',
        "provisional fingerprint is unknown")
    text = _replace_once(text, '''        let group = null;
        if (info.fp || info.dHash) {
          for (let gi = 0; gi < groups.length; gi += 1) {
            const candidate = groups[gi];
            const perceptualComparable = !!(info.dHash && info.aHash && candidate.dHash && candidate.aHash);
            const perceptualSource = !!(perceptualComparable &&
              dHashDistance(candidate.dHash, info.dHash) <= 2 &&
              dHashDistance(candidate.aHash, info.aHash) <= 4);
            // The strict fingerprint is only 32-bit for WPS compatibility.
            // When perceptual data exists, cross-check it so an accidental or
            // deliberately constructed FNV collision cannot force a merge.
            const exactSource = !!(info.fp && candidate.contentFp && info.fp === candidate.contentFp &&
              (!perceptualComparable || perceptualSource));
            if (!exactSource && !perceptualSource) continue;
            const windows = candidate._cropWindows || [];
            if (windows.every(function (existing) { return cropWindowCompatible(existing, cropWindow); })) {
              group = candidate;
              break;
            }
          }
        }
        if (!group && (info.fp || info.dHash)) {
          const sourceKey = info.dHash ? "d:" + info.dHash : "f:" + info.fp;''', '''        let group = null;
        // Similarity hashes cannot prove source equality. A destructive batch
        // group requires an exact SHA-256 identity of verified original media.
        // Missing identities and legacy metadata fail closed as separate rows.
        const strongFp = /^sha256:[a-f0-9]{64}$/i.test(String(info.fp || ""))
          ? String(info.fp).toLowerCase() : "";
        if (strongFp) {
          for (let gi = 0; gi < groups.length; gi += 1) {
            const candidate = groups[gi];
            if (candidate.contentFp !== strongFp) continue;
            const windows = candidate._cropWindows || [];
            if (windows.every(function (existing) { return cropWindowCompatible(existing, cropWindow); })) {
              group = candidate;
              break;
            }
          }
        }
        if (!group && strongFp) {
          const sourceKey = "f:" + strongFp;''', "strong source equality")
    text = _replace_once(text, '            contentFp: info.fp,', '            contentFp: strongFp,', "store canonical strong identity")
    return text
