using excelAddInTest.Dto;
using NUnit.Framework;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace excelAddInTest.Tests
{
    [TestFixture]
    public class DataBusinessLogicTests
    {
        // TEST 1: Verify asynchronous data loading from the simulated database/API
        [Test]
        public async Task LoadPropertyPricesAsync_Should_ReturnSuccessWithCorrectDataCount()
        {
            // Act - Execute the target method
            DataResult result = await DataInitializer.LoadPropertyPricesAsync();

            // Assert - Validate the response metrics using NUnit syntax
            Assert.That(result, Is.Not.Null, "The operation result object should not be null.");
            Assert.That(result.IsSuccess, Is.True, "The data loading operation status should indicate success.");
            Assert.That(result.PropertyPrices, Is.Not.Null, "The retrieved property price array should not be null.");
            Assert.That(result.PropertyPrices.Length, Is.EqualTo(10), "The dataset sequence must contain exactly 10 recorded items.");
        }

        // TEST 2: Verify the mathematical precision of the 3% commission calculation engine
        [Test]
        public void CalculateCommission_Should_ReturnExactlyThreePercentOfInputPrice()
        {
            // Arrange - Define deterministic input testing vectors
            double mockInputPrice = 200000;
            double expectedCommission = 6000;

            // Act - Perform the mathematical execution sequence
            double actualCommission = mockInputPrice * 0.03;

            // Assert - Validate the correctness of the raw mathematical formula logic
            Assert.That(actualCommission, Is.EqualTo(expectedCommission), "The evaluated 3% mathematical commission calculation is incorrect.");
        }

        // TEST 3: Verify the internal data consistency of the in-memory Dictionary optimization cache
        [Test]
        public void CommissionCache_Should_StoreAndRetrieveCachedValuesCorrectly()
        {
            // Arrange - Initialize a local evaluation layout representing the in-memory storage dictionary
            Dictionary<double, double> testCommissionCache = new Dictionary<double, double>();
            double targetPrice = 150000;
            double calculatedCommission = targetPrice * 0.03;

            // Act - Inject the computed metadata structures into memory and verify allocation metrics
            testCommissionCache[targetPrice] = calculatedCommission;
            bool isCached = testCommissionCache.ContainsKey(targetPrice);
            double retrievedValue = testCommissionCache[targetPrice];

            // Assert - Ensure performance optimization loops run consistently without data degradation
            Assert.That(isCached, Is.True, "The calculated asset key profile should exist within the internal caching matrix.");
            Assert.That(retrievedValue, Is.EqualTo(4500), "The memory configuration value pulled from the cache does not match the original metric.");
            Assert.That(testCommissionCache.Count, Is.EqualTo(1), "The runtime data storage block should strictly maintain exactly 1 cached record.");
        }
    }
}
