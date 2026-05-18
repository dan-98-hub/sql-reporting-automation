SELECT patient_email, COUNT(*)
FROM patients
GROUP BY patient_email
HAVING COUNT(*) > 1;
