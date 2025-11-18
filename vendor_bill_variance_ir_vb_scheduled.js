/**
 * @NApiVersion 2.1
 * @NScriptType ScheduledScript
 * @NModuleScope SameAccount
 */
define(['N/search', 'N/record', 'N/log', 'N/runtime'],
    /**
     * @param {search} search
     * @param {record} record
     * @param {log} log
     * @param {runtime} runtime
     */
    function (search, record, log, runtime) {

        /**
      * Executes the scheduled script
      * @param {Object} context
      */
        function execute(context) {
            log.audit('Script Start', 'Item Receipt Variance Update - Scheduled');

            var scriptObj = runtime.getCurrentScript();
            var initialUnits = scriptObj.getRemainingUsage();

            // Fixed thresholds
            var MIN_VARIANCE = 0.01;

            var results = {
                totalFound: 0,
                processed: 0,
                successCount: 0,
                errorCount: 0,
                closedPeriodCount: 0,
                errors: [],
                updated: [],
                closedPeriod: [],
                skipped: []
            };

            try {
                // Get variance pairs
                var variancePairs = getVariancePairs(MIN_VARIANCE);
                results.totalFound = variancePairs.length;

                log.audit('Variances Found', 'Total: ' + variancePairs.length);

                if (variancePairs.length === 0) {
                    log.audit('No Variances', 'No IR/VB rate variances found');
                    return;
                }

                log.audit('Processing Records', 'Processing ' + variancePairs.length + ' variances');

                // Process each variance
                variancePairs.forEach(function (pair) {
                    // Check governance
                    var remainingUsage = scriptObj.getRemainingUsage();
                    if (remainingUsage < 100) {
                        log.audit('Governance Limit', 'Stopping - only ' + remainingUsage + ' units remaining');
                        results.skipped.push({
                            reason: 'Governance limit reached',
                            count: variancePairs.length - results.processed
                        });
                        return; // Exit forEach
                    }

                    results.processed++;

                    try {
                        log.debug('Processing Variance', {
                            irNumber: pair.ir_number,
                            itemName: pair.item_name,
                            oldRate: pair.ir_rate,
                            newRate: pair.vb_rate,
                            variance: pair.vb_rate - pair.ir_rate
                        });

                        // Update the IR line rate
                        updateItemReceiptLineRate(pair.ir_id, pair.ir_line_id, pair.item_id, pair.vb_rate);

                        results.successCount++;
                        results.updated.push({
                            irNumber: pair.ir_number,
                            irId: pair.ir_id,
                            itemName: pair.item_name,
                            oldRate: pair.ir_rate,
                            newRate: pair.vb_rate,
                            variance: pair.vb_rate - pair.ir_rate,
                            vbNumber: pair.vb_number
                        });

                        log.audit('IR Updated', {
                            irNumber: pair.ir_number,
                            itemName: pair.item_name,
                            oldRate: pair.ir_rate,
                            newRate: pair.vb_rate
                        });

                    } catch (e) {
                        var errorMessage = e.message || e.toString();

                        // Check if this is a closed period error (expected behavior)
                        if (errorMessage.indexOf('closed period') !== -1) {
                            results.closedPeriodCount++;
                            results.closedPeriod.push({
                                irNumber: pair.ir_number,
                                irId: pair.ir_id,
                                itemName: pair.item_name,
                                vbNumber: pair.vb_number,
                                oldRate: pair.ir_rate,
                                newRate: pair.vb_rate,
                                variance: pair.vb_rate - pair.ir_rate,
                                reason: 'Period is closed'
                            });

                            log.audit('IR Skipped - Closed Period', {
                                irNumber: pair.ir_number,
                                itemName: pair.item_name,
                                oldRate: pair.ir_rate,
                                newRate: pair.vb_rate,
                                variance: pair.vb_rate - pair.ir_rate
                            });

                        } else {
                            // Actual errors (unexpected)
                            results.errorCount++;

                            // Simplify error message for other cases
                            if (errorMessage.indexOf('No lines found') !== -1) {
                                errorMessage = 'Item not found on Item Receipt';
                            } else if (errorMessage.indexOf('governance') !== -1 || errorMessage.indexOf('Usage Limit') !== -1) {
                                errorMessage = 'Script usage limit exceeded';
                            }

                            results.errors.push({
                                irNumber: pair.ir_number,
                                irId: pair.ir_id,
                                itemName: pair.item_name,
                                vbNumber: pair.vb_number,
                                error: errorMessage,
                                fullError: e.message || e.toString()
                            });

                            log.error('IR Update Failed', {
                                irNumber: pair.ir_number,
                                itemName: pair.item_name,
                                error: errorMessage,
                                fullError: e.message || e.toString()
                            });
                        }
                    }
                });

                // Calculate usage
                var usedUnits = initialUnits - scriptObj.getRemainingUsage();

                // Log final summary
                log.audit('Script Complete', {
                    totalFound: results.totalFound,
                    processed: results.processed,
                    successful: results.successCount,
                    closedPeriod: results.closedPeriodCount,
                    failed: results.errorCount,
                    skippedCount: results.skipped.length,
                    usageUnits: usedUnits
                });

                // Log successful updates detail
                if (results.updated.length > 0) {
                    log.audit('Successfully Updated', JSON.stringify(results.updated));
                }

                // Log closed period records (expected - not errors)
                if (results.closedPeriod.length > 0) {
                    log.audit('Skipped - Closed Period', JSON.stringify(results.closedPeriod));
                }

                // Log actual errors detail (unexpected issues)
                if (results.errors.length > 0) {
                    log.audit('Failed Updates', JSON.stringify(results.errors));
                }

                // Log skipped detail
                if (results.skipped.length > 0) {
                    log.audit('Skipped Records', JSON.stringify(results.skipped));
                }

            } catch (e) {
                log.error('Script Error', {
                    error: e.message || e.toString(),
                    stack: e.stack
                });
                throw e;
            }
        }

        /**
         * Updates a specific line on an Item Receipt with new rate
         * @param {string} irId - Item Receipt internal ID
         * @param {string} irLineId - IR Line unique ID
         * @param {string} itemId - Item internal ID
         * @param {number} newRate - New rate from vendor bill
         */
        function updateItemReceiptLineRate(irId, irLineId, itemId, newRate) {
            log.debug('Updating IR Line', {
                irId: irId,
                itemId: itemId,
                newRate: newRate
            });

            // Load the Item Receipt in standard mode
            var irRecord = record.load({
                type: record.Type.ITEM_RECEIPT,
                id: irId,
                isDynamic: false
            });

            // Find all matching lines by item ID and update them
            var lineCount = irRecord.getLineCount({ sublistId: 'item' });
            var linesUpdated = 0;

            for (var i = 0; i < lineCount; i++) {
                var currentItem = irRecord.getSublistValue({
                    sublistId: 'item',
                    fieldId: 'item',
                    line: i
                });

                // Match by item ID - update all instances of this item
                if (currentItem && currentItem.toString() === itemId.toString()) {
                    var oldRate = irRecord.getSublistValue({
                        sublistId: 'item',
                        fieldId: 'rate',
                        line: i
                    });

                    // Update the rate
                    irRecord.setSublistValue({
                        sublistId: 'item',
                        fieldId: 'rate',
                        line: i,
                        value: parseFloat(newRate)
                    });

                    linesUpdated++;
                    log.debug('Line Updated', {
                        line: i,
                        item: currentItem,
                        oldRate: oldRate,
                        newRate: newRate
                    });
                }
            }

            if (linesUpdated === 0) {
                throw new Error('No lines found with Item ID ' + itemId + ' on IR ' + irId);
            }

            // Save the record
            var savedId = irRecord.save({
                enableSourcing: false,
                ignoreMandatoryFields: true
            });

            log.audit('IR Updated Successfully', {
                irId: savedId,
                itemId: itemId,
                linesUpdated: linesUpdated,
                newRate: newRate
            });

            return savedId;
        }

        /**
         * Gets variance pairs by querying and matching oldest IR to oldest VB
         * @param {number} minVariance - Minimum variance to include (default 0.01)
         * @returns {Array} Array of variance pair objects
         */
        function getVariancePairs(minVariance) {
            var rawResults = searchIRVBVariances();
            var poLineGroups = groupByPOLine(rawResults);
            var variancePairs = createVariancePairs(poLineGroups, minVariance);

            log.debug('Variance Pairs Created', 'Total pairs: ' + variancePairs.length);

            return variancePairs;
        }

        /**
         * Searches for IR/VB rate variances
         * @returns {Array} Raw search results
         */
        function searchIRVBVariances() {
            var varianceSearch = search.create({
                type: search.Type.TRANSACTION,
                filters: [
                    ['type', 'anyof', 'PurchOrd'],
                    'AND',
                    ['mainline', 'is', 'F'],
                    'AND',
                    ['billingtransaction.quantity', 'greaterthan', '0'],
                    'AND',
                    ['fulfillingtransaction.quantity', 'greaterthan', '0'],
                    'AND',
                    ['formulanumeric: NVL({fulfillingtransaction.rate},0)-NVL({billingtransaction.amount}/{billingtransaction.quantity},0)', 'notequalto', '0']
                ],
                columns: [
                    search.createColumn({ name: 'internalid', label: 'PO ID' }),
                    search.createColumn({ name: 'tranid', label: 'PO Number' }),
                    search.createColumn({ name: 'trandate', label: 'PO Date', sort: search.Sort.ASC }),
                    search.createColumn({ name: 'entity', label: 'Vendor ID' }),
                    search.createColumn({ name: 'entityid', join: 'vendor', label: 'Vendor Name' }),
                    search.createColumn({ name: 'lineuniquekey', label: 'PO Line ID' }),
                    search.createColumn({ name: 'item', label: 'Item ID' }),
                    search.createColumn({ name: 'itemid', join: 'item', label: 'Item Number' }),
                    search.createColumn({ name: 'displayname', join: 'item', label: 'Item Name' }),
                    search.createColumn({ name: 'rate', label: 'PO Rate' }),
                    // Item Receipt columns
                    search.createColumn({ name: 'internalid', join: 'fulfillingtransaction', label: 'IR ID' }),
                    search.createColumn({ name: 'tranid', join: 'fulfillingtransaction', label: 'IR Number' }),
                    search.createColumn({ name: 'trandate', join: 'fulfillingtransaction', label: 'IR Date' }),
                    search.createColumn({ name: 'lineuniquekey', join: 'fulfillingtransaction', label: 'IR Line ID' }),
                    search.createColumn({ name: 'quantity', join: 'fulfillingtransaction', label: 'IR Quantity' }),
                    search.createColumn({ name: 'rate', join: 'fulfillingtransaction', label: 'IR Rate' }),
                    // Vendor Bill columns
                    search.createColumn({ name: 'internalid', join: 'billingtransaction', label: 'VB ID' }),
                    search.createColumn({ name: 'tranid', join: 'billingtransaction', label: 'VB Number' }),
                    search.createColumn({ name: 'trandate', join: 'billingtransaction', label: 'VB Date' }),
                    search.createColumn({ name: 'lineuniquekey', join: 'billingtransaction', label: 'VB Line ID' }),
                    search.createColumn({ name: 'quantity', join: 'billingtransaction', label: 'VB Quantity' }),
                    search.createColumn({ name: 'rate', join: 'billingtransaction', label: 'VB Rate' })
                ]
            });

            var results = [];
            varianceSearch.run().each(function (result) {
                var itemName = result.getText({ name: 'item' }) || result.getValue({ name: 'displayname', join: 'item' }) || '';

                results.push({
                    po_id: result.getValue({ name: 'internalid' }),
                    po_number: result.getValue({ name: 'tranid' }),
                    po_date: result.getValue({ name: 'trandate' }),
                    vendor_name: result.getValue({ name: 'entityid', join: 'vendor' }),
                    po_line_id: result.getValue({ name: 'lineuniquekey' }),
                    item_id: result.getValue({ name: 'item' }),
                    item_name: itemName,
                    ir_id: result.getValue({ name: 'internalid', join: 'fulfillingtransaction' }),
                    ir_number: result.getValue({ name: 'tranid', join: 'fulfillingtransaction' }),
                    ir_date: result.getValue({ name: 'trandate', join: 'fulfillingtransaction' }),
                    ir_line_id: result.getValue({ name: 'lineuniquekey', join: 'fulfillingtransaction' }),
                    ir_quantity: result.getValue({ name: 'quantity', join: 'fulfillingtransaction' }),
                    ir_rate: result.getValue({ name: 'rate', join: 'fulfillingtransaction' }),
                    vb_id: result.getValue({ name: 'internalid', join: 'billingtransaction' }),
                    vb_number: result.getValue({ name: 'tranid', join: 'billingtransaction' }),
                    vb_date: result.getValue({ name: 'trandate', join: 'billingtransaction' }),
                    vb_line_id: result.getValue({ name: 'lineuniquekey', join: 'billingtransaction' }),
                    vb_quantity: result.getValue({ name: 'quantity', join: 'billingtransaction' }),
                    vb_rate: result.getValue({ name: 'rate', join: 'billingtransaction' })
                });
                return true;
            });

            log.debug('Search Results', 'Total rows: ' + results.length);
            return results;
        }

        /**
         * Groups raw query results by PO Line ID
         * @param {Array} rawResults - Raw search results
         * @returns {Object} Grouped results by PO line
         */
        function groupByPOLine(rawResults) {
            var groups = {};

            rawResults.forEach(function (row) {
                var poLineKey = row.po_line_id;

                if (!groups[poLineKey]) {
                    groups[poLineKey] = {
                        poInfo: {
                            po_id: row.po_id,
                            po_number: row.po_number,
                            po_date: row.po_date,
                            vendor_name: row.vendor_name,
                            item_id: row.item_id,
                            item_name: row.item_name
                        },
                        itemReceipts: [],
                        vendorBills: []
                    };
                }

                // Add Item Receipt if not already added
                var irExists = groups[poLineKey].itemReceipts.some(function (ir) {
                    return ir.ir_line_id === row.ir_line_id;
                });
                if (!irExists) {
                    groups[poLineKey].itemReceipts.push({
                        ir_id: row.ir_id,
                        ir_number: row.ir_number,
                        ir_date: row.ir_date,
                        ir_line_id: row.ir_line_id,
                        ir_quantity: parseFloat(row.ir_quantity),
                        ir_rate: parseFloat(row.ir_rate)
                    });
                }

                // Add Vendor Bill if not already added
                var vbExists = groups[poLineKey].vendorBills.some(function (vb) {
                    return vb.vb_line_id === row.vb_line_id;
                });
                if (!vbExists) {
                    groups[poLineKey].vendorBills.push({
                        vb_id: row.vb_id,
                        vb_number: row.vb_number,
                        vb_date: row.vb_date,
                        vb_line_id: row.vb_line_id,
                        vb_quantity: parseFloat(row.vb_quantity),
                        vb_rate: parseFloat(row.vb_rate)
                    });
                }
            });

            return groups;
        }

        /**
         * Creates variance pairs by matching oldest IR to oldest VB
         * @param {Object} poLineGroups - Grouped results
         * @param {number} minVariance - Minimum variance to include
         * @returns {Array} Array of variance pair objects
         */
        function createVariancePairs(poLineGroups, minVariance) {
            var pairs = [];
            var threshold = minVariance || 0.01;

            Object.keys(poLineGroups).forEach(function (poLineKey) {
                var group = poLineGroups[poLineKey];

                // Sort by date (oldest first)
                group.itemReceipts.sort(function (a, b) {
                    return new Date(a.ir_date) - new Date(b.ir_date);
                });
                group.vendorBills.sort(function (a, b) {
                    return new Date(a.vb_date) - new Date(b.vb_date);
                });

                // Create 1:1 pairs (oldest to oldest)
                var maxPairs = Math.max(group.itemReceipts.length, group.vendorBills.length);

                for (var i = 0; i < maxPairs; i++) {
                    var ir = group.itemReceipts[i];
                    var vb = group.vendorBills[i];

                    if (ir && vb) {
                        var variance = vb.vb_rate - ir.ir_rate;

                        if (Math.abs(variance) >= threshold) {
                            pairs.push({
                                po_id: group.poInfo.po_id,
                                po_number: group.poInfo.po_number,
                                po_date: group.poInfo.po_date,
                                vendor_name: group.poInfo.vendor_name,
                                item_id: group.poInfo.item_id,
                                item_name: group.poInfo.item_name,
                                ir_id: ir.ir_id,
                                ir_number: ir.ir_number,
                                ir_date: ir.ir_date,
                                ir_line_id: ir.ir_line_id,
                                ir_quantity: ir.ir_quantity,
                                ir_rate: ir.ir_rate,
                                vb_id: vb.vb_id,
                                vb_number: vb.vb_number,
                                vb_date: vb.vb_date,
                                vb_line_id: vb.vb_line_id,
                                vb_quantity: vb.vb_quantity,
                                vb_rate: vb.vb_rate
                            });
                        }
                    }
                }
            });

            return pairs;
        }

        return {
            execute: execute
        };
    });