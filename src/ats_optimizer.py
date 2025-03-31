import re
from typing import List, Dict
import spacy
from docx import Document

class ATSOptimizer:
    def __init__(self):
        # Load English language model for NLP
        try:
            self.nlp = spacy.load("en_core_web_sm")
        except OSError:
            print("Downloading language model...")
            spacy.cli.download("en_core_web_sm")
            self.nlp = spacy.load("en_core_web_sm")

    def optimize_text(self, text: str) -> str:
        """Optimize text for ATS by improving formatting and removing problematic elements."""
        # Remove special characters and extra whitespace
        text = re.sub(r'[^\w\s.,;:()]', '', text)
        text = re.sub(r'\s+', ' ', text)
        
        # Convert to proper case for important terms
        doc = self.nlp(text)
        optimized_text = ""
        
        for token in doc:
            if token.pos_ in ['NOUN', 'PROPN', 'VERB']:
                optimized_text += token.text.capitalize() + " "
            else:
                optimized_text += token.text + " "
        
        return optimized_text.strip()

    def extract_keywords(self, text: str) -> List[str]:
        """Extract important keywords from text."""
        doc = self.nlp(text)
        keywords = []
        
        # Extract nouns, verbs, and adjectives
        for token in doc:
            if token.pos_ in ['NOUN', 'VERB', 'ADJ']:
                keywords.append(token.text.lower())
        
        return list(set(keywords))  # Remove duplicates

    def optimize_document(self, doc: Document) -> Document:
        """Optimize a Word document for ATS."""
        for paragraph in doc.paragraphs:
            # Remove any tables or complex formatting
            if paragraph._p.getparent().tag.endswith('tbl'):
                continue
                
            # Optimize text content
            optimized_text = self.optimize_text(paragraph.text)
            paragraph.text = optimized_text
            
            # Ensure proper font and size
            for run in paragraph.runs:
                run.font.name = 'Arial'
                run.font.size = 110000  # 11pt
        
        return doc

    def analyze_job_description(self, job_description: str) -> Dict[str, List[str]]:
        """Analyze a job description to extract important keywords and requirements."""
        doc = self.nlp(job_description)
        
        analysis = {
            'required_skills': [],
            'preferred_skills': [],
            'experience_level': [],
            'education_requirements': [],
            'keywords': []
        }
        
        # Extract keywords
        analysis['keywords'] = self.extract_keywords(job_description)
        
        # Look for required skills (usually after "required:", "must have:", etc.)
        required_patterns = [
            r'required:.*?(?=\n|$)',
            r'must have:.*?(?=\n|$)',
            r'requirements:.*?(?=\n|$)'
        ]
        
        for pattern in required_patterns:
            matches = re.finditer(pattern, job_description, re.IGNORECASE)
            for match in matches:
                skills_text = match.group(0).split(':', 1)[1].strip()
                analysis['required_skills'].extend(self.extract_keywords(skills_text))
        
        # Look for preferred skills
        preferred_patterns = [
            r'preferred:.*?(?=\n|$)',
            r'nice to have:.*?(?=\n|$)',
            r'bonus:.*?(?=\n|$)'
        ]
        
        for pattern in preferred_patterns:
            matches = re.finditer(pattern, job_description, re.IGNORECASE)
            for match in matches:
                skills_text = match.group(0).split(':', 1)[1].strip()
                analysis['preferred_skills'].extend(self.extract_keywords(skills_text))
        
        # Look for experience level
        experience_patterns = [
            r'(\d+)\+?\s*(?:years?|yrs?)\s*of\s*experience',
            r'experience\s*level:\s*(entry|mid|senior|lead|principal)',
            r'(entry|mid|senior|lead|principal)\s*level'
        ]
        
        for pattern in experience_patterns:
            matches = re.finditer(pattern, job_description, re.IGNORECASE)
            for match in matches:
                analysis['experience_level'].append(match.group(0))
        
        # Look for education requirements
        education_patterns = [
            r'bachelor\'s|master\'s|phd|degree',
            r'bs|ms|phd',
            r'education:\s*.*?(?=\n|$)'
        ]
        
        for pattern in education_patterns:
            matches = re.finditer(pattern, job_description, re.IGNORECASE)
            for match in matches:
                analysis['education_requirements'].append(match.group(0))
        
        return analysis

    def get_optimization_suggestions(self, resume_text: str, job_description: str) -> List[str]:
        """Generate suggestions for optimizing a resume based on a job description."""
        suggestions = []
        
        # Analyze job description
        job_analysis = self.analyze_job_description(job_description)
        
        # Extract resume keywords
        resume_keywords = self.extract_keywords(resume_text)
        
        # Check for missing required skills
        missing_required = set(job_analysis['required_skills']) - set(resume_keywords)
        if missing_required:
            suggestions.append(f"Consider adding these required skills: {', '.join(missing_required)}")
        
        # Check for missing preferred skills
        missing_preferred = set(job_analysis['preferred_skills']) - set(resume_keywords)
        if missing_preferred:
            suggestions.append(f"Consider adding these preferred skills: {', '.join(missing_preferred)}")
        
        # Check for experience level match
        if job_analysis['experience_level']:
            suggestions.append(f"Ensure your experience matches the required level: {', '.join(job_analysis['experience_level'])}")
        
        # Check for education requirements
        if job_analysis['education_requirements']:
            suggestions.append(f"Verify you meet the education requirements: {', '.join(job_analysis['education_requirements'])}")
        
        return suggestions 