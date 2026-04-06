function isLaserLProduct(prodNo) {
    return String(prodNo || '').trim().toUpperCase().endsWith('L');
}

function isR1100Operation(prodNo) {
    return String(prodNo || '').trim().toUpperCase() === 'R1100';
}

function isGloballyExcludedProdNo(prodNo) {
    return String(prodNo || '').trim().toUpperCase() === 'R1090';
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
    isLaserEagleOperator,
    shouldDoubleR1100Operation,
    adjustOperationLinePricing
};
