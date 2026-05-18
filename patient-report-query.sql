SELECT patient_id, appointment_date, provider_name
FROM appointments
WHERE appointment_date >= '2026-01-01'
ORDER BY appointment_date DESC;
