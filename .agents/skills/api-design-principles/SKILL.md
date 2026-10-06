---
name: api-design-principles
description: >-
  Provides comprehensive standards and best practices for RESTful API design, webhook architecture, and endpoint integration. Use when designing, building, or reviewing HTTP APIs, webhook receivers, or Google Apps Script doGet/doPost endpoints connecting forms, external services, or frontend applications to backend sheets and databases.
---

# API Design & Webhook Engineering Principles

This skill defines industry-standard patterns for building robust, secure, and developer-friendly REST APIs and webhook ingestion pipelines, including serverless endpoints hosted via Google Apps Script Web Apps (`doGet` / `doPost`).

---

## 1. RESTful Modeling & HTTP Semantics

### Standard HTTP Methods
- **`GET`**: Retrieve resources. Must be **safe** (no side effects) and **idempotent**.
- **`POST`**: Create a new resource or execute a non-idempotent action (e.g. process payment, submit form).
- **`PUT`**: Replace an existing resource entirely (idempotent).
- **`PATCH`**: Partially update specific fields of an existing resource.
- **`DELETE`**: Remove a resource (idempotent).

### URL Naming Conventions
- Use plural nouns for resource collections: `/api/v1/transactions`, `/api/v1/statements`.
- Express hierarchy naturally: `/api/v1/statements/{id}/transactions`.
- Avoid verbs in endpoints: prefer `POST /api/v1/transactions/{id}/reconcile` over `/api/v1/doReconcileTransaction`.

---

## 2. Uniform JSON Response Schema

All responses must adhere to an unambiguous, consistent envelope format:

### Success Response
```json
{
  "success": true,
  "data": {
    "id": "tx_98124",
    "amount": 450.00,
    "currency": "EGP",
    "status": "settled",
    "createdAt": "2026-10-06T19:28:00Z"
  },
  "meta": {
    "timestamp": 1791244080000,
    "version": "1.0"
  }
}
```

### Error Response
```json
{
  "success": false,
  "error": {
    "code": "VALIDATION_FAILED",
    "message": "The transaction amount must be greater than zero.",
    "details": [
      {
        "field": "amount",
        "issue": "Must be positive number"
      }
    ]
  },
  "meta": {
    "timestamp": 1791244080000
  }
}
```

---

## 3. Webhook Architecture & Idempotency

Webhooks are often retried automatically by third-party services (Stripe, Fawry, GitHub, form builders) if an acknowledgment is delayed.

### Idempotency Key Handling
1. Require an `Idempotency-Key` or unique transaction reference in every payload or header.
2. Check if the key has already been processed (e.g. in `PropertiesService` or database):
   - If already processed: return the cached successful response immediately without re-executing business logic.
   - If new: mark the key as processing, execute the action, and record the result.

### Fast Acknowledgment
- Return an HTTP `200` or `202 Accepted` response within **3 seconds** of receiving the webhook.
- Avoid performing heavy PDF OCR, complex reconciliations, or multi-tab sheet updates synchronously inside the webhook request. If necessary, queue the job and trigger execution asynchronously.

### Authentication & Signature Verification
- Protect endpoints with a shared bearer secret or HMAC signature:
  ```javascript
  function authenticateRequest(e, expectedSecret) {
    const authHeader = e.parameter?.token || e.headers?.["Authorization"] || "";
    const cleanToken = authHeader.replace(/^Bearer\s+/i, "");
    return cleanToken === expectedSecret;
  }
  ```

---

## 4. Google Apps Script Web App Implementation (`doGet` / `doPost`)

When exposing Google Sheets as a backend API using Apps Script:

### Complete Endpoint Template
```javascript
function doPost(e) {
  const lock = LockService.getScriptLock();
  // Protect against concurrent row append collisions
  if (!lock.tryLock(10000)) {
    return createJsonResponse({
      success: false,
      error: { code: "SERVER_BUSY", message: "Concurrent request limit reached. Please retry." }
    }, 429);
  }

  try {
    if (!e || !e.postData || !e.postData.contents) {
      return createJsonResponse({
        success: false,
        error: { code: "EMPTY_PAYLOAD", message: "No request body provided." }
      }, 400);
    }

    const body = JSON.parse(e.postData.contents);

    // Validate required fields
    if (!body.amount || isNaN(parseFloat(body.amount)) || parseFloat(body.amount) <= 0) {
      return createJsonResponse({
        success: false,
        error: { code: "INVALID_AMOUNT", message: "A positive numeric amount is required." }
      }, 422);
    }

    // Execute business logic (e.g. append to sheet)
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Transactions");
    sheet.appendRow([
      new Date(),
      body.date || new Date(),
      body.description || "API Submission",
      body.payer || "Mido",
      parseFloat(body.amount)
    ]);

    return createJsonResponse({
      success: true,
      data: {
        message: "Transaction recorded successfully.",
        recordedAmount: parseFloat(body.amount)
      }
    }, 201);

  } catch (err) {
    return createJsonResponse({
      success: false,
      error: { code: "INTERNAL_ERROR", message: err?.message || String(err) }
    }, 500);
  } finally {
    lock.releaseLock();
  }
}

function createJsonResponse(payload, statusCode) {
  // Apps Script ContentService serves plain text with JSON mimeType
  return ContentService.createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}
```

---

## 5. HTTP Status Code Guidelines

| Status Code | Meaning | When to Use |
| :--- | :--- | :--- |
| **`200 OK`** | Success | Successful `GET` or non-creating `POST` |
| **`201 Created`** | Resource Created | New transaction or entry successfully inserted |
| **`202 Accepted`** | Request Queued | Asynchronous task accepted for batch processing |
| **`400 Bad Request`** | Malformed Input | Missing required JSON body or corrupt syntax |
| **`401 Unauthorized`** | Missing Auth | Missing or incorrect API token/bearer key |
| **`403 Forbidden`** | Insufficient Perms | Valid token, but forbidden to access resource |
| **`404 Not Found`** | Missing Resource | Statement ID or sheet record does not exist |
| **`409 Conflict`** | State Conflict | Duplicate transaction reference detected |
| **`422 Unprocessable`** | Semantic Error | Valid JSON, but fails validation (e.g. negative amount) |
| **`429 Too Many Req`** | Rate Limit | Concurrency lock timeout or quota exceeded |
| **`500 Internal Error`**| Server Exception | Uncaught runtime error |
