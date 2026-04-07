function normalizeProdNo(prodNo) {
    return String(prodNo || '').trim().toUpperCase();
}

function isLaserLProduct(prodNo) {
    return normalizeProdNo(prodNo).endsWith('L');
}

function isR1100Operation(prodNo) {
    return normalizeProdNo(prodNo) === 'R1100';
}

function isGloballyExcludedProdNo(prodNo) {
    return normalizeProdNo(prodNo) === 'R1090';
}

function isExcludedOperationProdNo(prodNo) {
    const normalized = normalizeProdNo(prodNo);
    return normalized === 'R1090' || normalized === 'R8200';
}

function isEstimatedOperationMinutesFallback(line) {
    if (!line) return false;

    const key = (line.ProdTp4 === null || line.ProdTp4 === undefined) ? 'NA' : String(line.ProdTp4);
    const normalizedKey = key === '3' ? '1' : key;
    const prodNoKey = normalizeProdNo(line.ProdNo);
    const noFin = Number(line.NoFin || 0);
    const noOrg = Number(line.NoOrg || 0);

    return normalizedKey === '1'
        && prodNoKey.startsWith('R')
        && !isExcludedOperationProdNo(prodNoKey)
        && noFin === 0
        && noOrg > 0;
}

function getEffectiveOperationMinutes(line) {
    if (!line) return 0;

    return isEstimatedOperationMinutesFallback(line)
        ? Number(line.NoOrg || 0)
        : Number(line.NoFin || 0);
}

function isLaserEagleOperator(hvemNm) {
    return String(hvemNm || '').trim().toUpperCase().includes('LASER EAGLE');
}

function shouldDoubleR1100Operation(prodNo, hvemNm, prodTp4) {
    const key = (prodTp4 === null || prodTp4 === undefined) ? 'NA' : String(prodTp4);
    return key === '1' && isR1100Operation(prodNo) && isLaserEagleOperator(hvemNm);
}

function adjustOperationLinePricing(line) {
    if (!line || !shouldDoubleR1100Operation(line.ProdNo, line.HvemNm, line.ProdTp4)) {
        return line;
    }

    const quantity = Number(line.NoFin || 0);
    const doubledDPrice = Number(line.DPrice || 0) * 2;
    const doubledCCstPr = Number(line.CCstPr || 0) * 2;
    const doubledLineCost = quantity > 0
        ? doubledCCstPr * quantity
        : Number(line.LineCost || 0) * 2;

    return {
        ...line,
        DPrice: doubledDPrice,
        CCstPr: doubledCCstPr,
        LineCost: doubledLineCost
    };
}

module.exports = {
    isLaserLProduct,
    isR1100Operation,
    isGloballyExcludedProdNo,
    isExcludedOperationProdNo,
    isEstimatedOperationMinutesFallback,
    getEffectiveOperationMinutes,
    isLaserEagleOperator,
    shouldDoubleR1100Operation,
    adjustOperationLinePricing
};
