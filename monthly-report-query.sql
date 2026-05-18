SELECT department, COUNT(*) AS total_appointments
FROM appointments
WHERE appointment_date BETWEEN '2020-01-01' AND '2021-01-31'
GROUP BY department;
