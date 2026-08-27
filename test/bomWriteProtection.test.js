const test = require('node:test');
const assert = require('node:assert/strict');
const { createBomService } = require('../services/bomService');

function createDiskCache() {
    return {
        get: () => null,
        set: () => {},
        list: () => [],
        del: () => {}
    };
}

function createSqlHarness(existingRows = []) {
    const state = {
        began: 0,
        committed: 0,
        rolledBack: 0,
        insertQueries: []
    };

    class Transaction {
        async begin() { state.began += 1; }
        async commit() { state.committed += 1; }
        async rollback() { state.rolledBack += 1; }
    }

    class Request {
        input() { return this; }
        async query(query) {
            state.insertQueries.push(query);
            return { recordset: [] };
        }
    }

    const pool = {
        request() {
            return {
                input() { return this; },
                async query() { return { recordset: existingRows }; }
            };
        }
    };

    return {
        state,
        pool,
        sql: {
            VarChar: () => 'VarChar',
            Int: 'Int',
            Float: 'Float',
            Transaction,
            Request
        }
    };
}

const productInput = {
    customerCode: '100',
    prodNoSuffix: '-TEST',
    descr: 'Test product'
};

test('readOnly blocks the execute command before any database access', async () => {
    let connectionCalls = 0;
    const harness = createSqlHarness();
    const service = createBomService({
        getConnection: async () => {
            connectionCalls += 1;
            return harness.pool;
        },
        sql: harness.sql,
        diskCache: createDiskCache(),
        getActiveProfile: () => ({ id: 'test', readOnly: true })
    });

    await assert.rejects(
        service.createProductsInVisma(productInput),
        error => error.statusCode === 403 && error.code === 'READ_ONLY_PROFILE'
    );
    assert.equal(connectionCalls, 0);
    assert.equal(harness.state.began, 0);
});

test('readOnly still permits preview and duplicate checking', async () => {
    const harness = createSqlHarness([{ ProdNo: '100-TEST', Descr: 'Existing' }]);
    const service = createBomService({
        getConnection: async () => harness.pool,
        sql: harness.sql,
        diskCache: createDiskCache(),
        getActiveProfile: () => ({ id: 'test', readOnly: true })
    });

    const preview = await service.previewCreateProducts(productInput);
    assert.deepEqual(preview.conflicts, ['100-TEST']);
    assert.equal(preview.canCreate, false);
});

test('writable profiles preserve the existing transaction and cache flow', async () => {
    const harness = createSqlHarness();
    const service = createBomService({
        getConnection: async () => harness.pool,
        sql: harness.sql,
        diskCache: createDiskCache(),
        getActiveProfile: () => ({ id: 'production', readOnly: false })
    });

    const result = await service.createProductsInVisma(productInput);

    assert.equal(result.created.length, 1);
    assert.equal(result.created[0].ProdNo, '100-TEST');
    assert.equal(harness.state.began, 1);
    assert.equal(harness.state.committed, 1);
    assert.equal(harness.state.rolledBack, 0);
    assert.equal(harness.state.insertQueries.length, 1);
    assert.match(harness.state.insertQueries[0], /INSERT INTO Prod/);
});

test('a failed Visma insert still rolls the existing transaction back', async () => {
    const harness = createSqlHarness();
    harness.sql.Request = class {
        input() { return this; }
        async query() { throw new Error('simulated insert failure'); }
    };
    const service = createBomService({
        getConnection: async () => harness.pool,
        sql: harness.sql,
        diskCache: createDiskCache(),
        getActiveProfile: () => ({ id: 'production', readOnly: false })
    });

    await assert.rejects(service.createProductsInVisma(productInput), /simulated insert failure/);
    assert.equal(harness.state.began, 1);
    assert.equal(harness.state.committed, 0);
    assert.equal(harness.state.rolledBack, 1);
});
