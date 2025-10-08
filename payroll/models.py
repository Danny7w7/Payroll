from django.db import models

# Create your models here.
# models.py
import uuid
from django.db import models
from django.utils import timezone
from datetime import timedelta

class PaymentToken(models.Model):
    token = models.UUIDField(default=uuid.uuid4, editable=False, unique=True)
    stripe_session_id = models.CharField(max_length=255, unique=True)
    
    is_used = models.BooleanField(default=False)
    is_paid = models.BooleanField(default=False)
    
    created_at = models.DateTimeField(auto_now_add=True)
    paid_at = models.DateTimeField(null=True, blank=True)
    used_at = models.DateTimeField(null=True, blank=True)
    expires_at = models.DateTimeField()
    
    customer_email = models.EmailField(blank=True, null=True)
    
    class Meta:
        ordering = ['-created_at']
    
    def __str__(self):
        return f"Token {self.token} - {'Usado' if self.is_used else 'Pagado' if self.is_paid else 'Pendiente'}"
    
    def save(self, *args, **kwargs):
        if not self.expires_at:
            # El token expira en 24 horas después del pago
            self.expires_at = timezone.now() + timedelta(hours=24)
        super().save(*args, **kwargs)
    
    def is_valid(self):
        """Verifica si el token es válido para usar"""
        return (
            self.is_paid and 
            not self.is_used and 
            timezone.now() < self.expires_at
        )
    
    def mark_as_used(self):
        """Marca el token como usado"""
        self.is_used = True
        self.used_at = timezone.now()
        self.save()