SELECT department, COUNT(*) AS total_appointments
FROM appointments
WHERE appointment_date BETWEEN '2026-01-01' AND '2026-01-31'
GROUP BY department;
