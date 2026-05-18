SELECT *
FROM patients
WHERE patient_phone IS NULL
OR patient_email IS NULL;
